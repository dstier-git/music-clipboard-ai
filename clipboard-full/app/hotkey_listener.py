import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

if __package__ is None or __package__ == "":
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.platform_utils import IS_MACOS, IS_WINDOWS, default_hotkey

try:
    if IS_MACOS:
        from pynput import keyboard as pynput_keyboard
        keyboard = None
    else:
        import keyboard
        pynput_keyboard = None
except ImportError:
    keyboard = None
    pynput_keyboard = None

try:
    import psutil
except ImportError:
    psutil = None

APP_SCRIPT = Path(__file__).resolve().parent / "musescore_extractor_gui.py"
REQUEST_FILE = Path(tempfile.gettempdir()) / "musescore_hotkey_request.txt"
HOTKEY = default_hotkey()


def _get_interpreter():
    interpreter = Path(sys.executable)
    if IS_WINDOWS:
        pythonw = interpreter.with_name("pythonw.exe")
        if pythonw.exists():
            return pythonw
    return interpreter


def _is_gui_running():
    if psutil is None:
        return False

    for proc in psutil.process_iter(["cmdline", "name"]):
        try:
            cmdline = proc.info.get("cmdline") or []
            name = proc.info.get("name") or ""
            for part in cmdline:
                if not part:
                    continue
                if APP_SCRIPT.name in os.path.basename(part):
                    return True
            if "python" in name.lower() and any(APP_SCRIPT.name in str(arg) for arg in cmdline):
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False


def _start_gui():
    interpreter = _get_interpreter()
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if IS_WINDOWS else 0
    stdout = subprocess.DEVNULL if IS_MACOS else None
    stderr = subprocess.DEVNULL if IS_MACOS else None
    subprocess.Popen(
        [str(interpreter), str(APP_SCRIPT), "--trigger-save-selection", "--disable-global-hotkey"],
        creationflags=creation_flags,
        stdout=stdout,
        stderr=stderr,
    )


def _signal_gui():
    try:
        REQUEST_FILE.parent.mkdir(parents=True, exist_ok=True)
        REQUEST_FILE.write_text(str(time.time()))
    except Exception:
        pass


def _on_hotkey():
    if _is_gui_running():
        _signal_gui()
    else:
        _start_gui()


def main():
    if IS_MACOS:
        if pynput_keyboard is None:
            raise SystemExit(
                "The 'pynput' library is required to run hotkey_listener.py on macOS. "
                "Install it with: pip install pynput"
            )
    else:
        if keyboard is None:
            raise SystemExit(
                "The 'keyboard' library is required to run hotkey_listener.py. "
                "Install it with: pip install keyboard"
            )

    if psutil is None:
        print(
            "Warning: psutil is not installed. The listener will always restart the GUI instead of signaling the running one."
        )

    try:
        REQUEST_FILE.parent.mkdir(parents=True, exist_ok=True)
        REQUEST_FILE.write_text(str(time.time()))
    except Exception:
        pass

    display_hotkey = HOTKEY.replace("cmd", "Cmd").replace("ctrl", "Ctrl").replace("alt", "Alt")
    print(f"Listening for {display_hotkey} -> launches {APP_SCRIPT.name}")

    if IS_MACOS:
        pynput_hotkey = (
            HOTKEY.replace("cmd", "<cmd>")
            .replace("ctrl", "<ctrl>")
            .replace("alt", "<alt>")
            .replace("shift", "<shift>")
        )
        with pynput_keyboard.GlobalHotKeys({pynput_hotkey: _on_hotkey}) as listener:
            listener.join()
    else:
        keyboard.add_hotkey(HOTKEY, _on_hotkey)
        keyboard.wait()


if __name__ == "__main__":
    main()
