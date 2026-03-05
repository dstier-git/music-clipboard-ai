import argparse
import json
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk

if __package__ is None or __package__ == "":
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.platform_utils import IS_MACOS, IS_WINDOWS, default_hotkey, output_dirs

# Try to import automation libraries
try:
    import pyautogui
    pyautogui.FAILSAFE = False
    PYAUTOGUI_AVAILABLE = True
except ImportError:
    PYAUTOGUI_AVAILABLE = False

try:
    from pywinauto import Application
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False

try:
    import win32con
    import win32gui
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

try:
    import keyboard
    KEYBOARD_AVAILABLE = True
except ImportError:
    KEYBOARD_AVAILABLE = False

if IS_MACOS:
    try:
        from pynput import keyboard as pynput_keyboard
        PYNPUT_AVAILABLE = True
    except ImportError:
        pynput_keyboard = None
        PYNPUT_AVAILABLE = False
else:
    pynput_keyboard = None
    PYNPUT_AVAILABLE = False

try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

CONFIG_FILE = Path(os.path.expanduser("~")) / ".musescore_pitch_extractor_prefs"
HOTKEY_REQUEST_FILE = Path(tempfile.gettempdir()) / "musescore_hotkey_request.txt"
WATCHED_SCORE_EXTENSIONS = (".mscx", ".mscz", ".mid", ".midi")
EXTRACTABLE_SCORE_EXTENSIONS = (".mscx", ".mscz")

OUTPUT_DIR, MIDI_OUTPUT_DIR = output_dirs()

GLOBAL_HOTKEY = default_hotkey()

PROGRAM_PROFILES = {
    "musescore": {
        "label": "MuseScore",
        "mac_process_names": ["mscore", "MuseScore 4", "MuseScore 3", "MuseScore"],
        "mac_app_names": ["MuseScore 4", "MuseScore 3", "MuseScore"],
        "mac_contains": ["musescore", "mscore"],
        "windows_process_keywords": ["musescore4.exe", "musescore", "mscore"],
        "windows_title_keywords": ["musescore 4", "musescore"],
        "default_hotkeys": {
            "mac": "cmd+shift+s",
            "windows": "ctrl+shift+s",
        },
    },
    "logic_pro": {
        "label": "Logic Pro",
        "mac_process_names": ["Logic Pro", "Logic Pro X"],
        "mac_app_names": ["Logic Pro", "Logic Pro X"],
        "mac_contains": ["logic pro", "logic"],
        "windows_process_keywords": ["logic", "logicpro"],
        "windows_title_keywords": ["logic pro", "logic"],
        "default_hotkeys": {
            "mac": "cmd+alt+e",
            "windows": "",
        },
    },
}

PROGRAM_ORDER = ["musescore", "logic_pro"]
HOTKEY_MODIFIER_ORDER = ["cmd", "ctrl", "alt", "shift"]
HOTKEY_MODIFIER_ALIASES = {
    "cmd": "cmd",
    "command": "cmd",
    "ctrl": "ctrl",
    "control": "ctrl",
    "alt": "alt",
    "option": "alt",
    "shift": "shift",
}

EXTRACTION_FUNCTION = None
MIDI_EXTRACTION_FUNCTION = None
try:
    from app.extract_pitches_with_position import extract_pitches_with_position_from_mscx

    EXTRACTION_FUNCTION = extract_pitches_with_position_from_mscx
    EXTRACTION_SCRIPT = "extract_pitches_with_position"
except ImportError:
    try:
        from app.extract_pitches import extract_pitches_from_mscx

        EXTRACTION_SCRIPT = "extract_pitches"

        def extract_pitches_with_position_from_mscx(file_path, output_file_path=None, debug=False):
            pitches = extract_pitches_from_mscx(file_path, output_file_path, debug)
            if pitches:
                return [(p, "N/A", None) for p in pitches]
            return None

        EXTRACTION_FUNCTION = extract_pitches_with_position_from_mscx
    except ImportError:
        EXTRACTION_FUNCTION = None
        EXTRACTION_SCRIPT = None

try:
    from app.extract_midi import extract_midi_from_mscx

    MIDI_EXTRACTION_FUNCTION = extract_midi_from_mscx
except ImportError:
    MIDI_EXTRACTION_FUNCTION = None


def _format_hotkey_label(hotkey):
    if not hotkey:
        return "Not configured"
    parts = []
    for part in hotkey.split("+"):
        if part == "cmd":
            parts.append("Cmd")
        elif part == "space":
            parts.append("Space")
        else:
            parts.append(part.capitalize())
    return "+".join(parts)



def _current_platform_key():
    return "mac" if IS_MACOS else "windows" if IS_WINDOWS else "other"


def _normalize_hotkey_value(raw_hotkey, platform_key):
    text = (raw_hotkey or "").strip().lower()
    if not text:
        return "", None

    tokens = [token.strip() for token in text.split("+") if token.strip()]
    if not tokens:
        return "", "Hotkey is empty."

    modifiers = []
    key_tokens = []
    for token in tokens:
        mapped = HOTKEY_MODIFIER_ALIASES.get(token, token)
        if mapped in HOTKEY_MODIFIER_ORDER:
            if mapped not in modifiers:
                modifiers.append(mapped)
        else:
            key_tokens.append(mapped)

    if len(key_tokens) != 1:
        return "", "Hotkey must contain exactly one non-modifier key."

    key = key_tokens[0]
    if key == "space":
        key = "space"
    elif len(key) != 1:
        return "", "Only single-character keys (or 'space') are supported."
    elif not key.isalnum():
        return "", "Only alphanumeric keys (or 'space') are supported."

    if platform_key == "windows" and "cmd" in modifiers:
        return "", "Windows hotkeys cannot use 'cmd'. Use ctrl/alt/shift."

    ordered_modifiers = [mod for mod in HOTKEY_MODIFIER_ORDER if mod in modifiers]
    normalized = "+".join(ordered_modifiers + [key])
    return normalized, None


def _split_normalized_hotkey(normalized_hotkey):
    parts = [part for part in (normalized_hotkey or "").split("+") if part]
    if not parts:
        return [], ""
    key = parts[-1]
    modifiers = parts[:-1]
    return modifiers, key


def _hotkey_to_windows_pywinauto(normalized_hotkey):
    modifiers, key = _split_normalized_hotkey(normalized_hotkey)
    prefix = ""
    if "ctrl" in modifiers:
        prefix += "^"
    if "alt" in modifiers:
        prefix += "%"
    if "shift" in modifiers:
        prefix += "+"

    if key == "space":
        key_token = "{SPACE}"
    elif key in ["+", "^", "%", "~", "(", ")", "{", "}"]:
        key_token = "{" + key + "}"
    else:
        key_token = key
    return prefix + key_token


def run_applescript(script):
    """Run an AppleScript command and return the result."""
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=10,
        )
        if result.returncode != 0:
            error_msg = result.stderr.strip() or "AppleScript returned non-zero exit code"
            return False, result.stdout.strip(), error_msg
        return True, result.stdout.strip(), result.stderr.strip()
    except subprocess.TimeoutExpired:
        return False, "", "Timeout"
    except Exception as e:
        return False, "", str(e)


def find_musescore_window_macos():
    """Find MuseScore window on macOS using AppleScript with multiple fallback methods."""
    scripts = [
        """
        tell application "System Events"
            try
                set museScoreProcess to first process whose name is "mscore"
                return name of museScoreProcess
            on error
                return ""
            end try
        end tell
        """,
        """
        tell application "System Events"
            try
                set museScoreProcess to first process whose name is "MuseScore 4"
                return name of museScoreProcess
            on error
                return ""
            end try
        end tell
        """,
        """
        tell application "System Events"
            try
                set museScoreProcess to first process whose name contains "MuseScore"
                return name of museScoreProcess
            on error
                return ""
            end try
        end tell
        """,
        """
        tell application "System Events"
            set processList to name of every process
            repeat with procName in processList
                if procName contains "MuseScore" or procName is "mscore" then
                    return procName
                end if
            end repeat
            return ""
        end tell
        """,
    ]

    for script in scripts:
        success, output, error = run_applescript(script)
        if success and output and output.strip():
            return True, output.strip(), ""

    if PSUTIL_AVAILABLE:
        try:
            for proc in psutil.process_iter(["pid", "name"]):
                try:
                    proc_name = proc.info.get("name") or ""
                    proc_lower = proc_name.lower()
                    if "musescore" in proc_lower or proc_lower == "mscore":
                        return True, proc_name, ""
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
        except Exception:
            pass

    return False, "", "MuseScore process not found"


def activate_musescore_window_macos():
    """Activate MuseScore window on macOS using AppleScript with multiple fallback methods."""
    scripts = [
        """
        tell application "System Events"
            try
                set museScoreProcess to first process whose name is "mscore"
                set frontmost of museScoreProcess to true
            on error
                return false
            end try
        end tell
        """,
        """
        tell application "MuseScore 4"
            activate
        end tell
        """,
        """
        tell application "MuseScore 3"
            activate
        end tell
        """,
        """
        tell application "MuseScore"
            activate
        end tell
        """,
        """
        tell application "System Events"
            try
                set museScoreProcess to first process whose name is "mscore"
                set frontmost of museScoreProcess to true
            on error
                try
                    set museScoreProcess to first process whose name is "MuseScore 4"
                    set frontmost of museScoreProcess to true
                on error
                    set museScoreProcess to first process whose name contains "MuseScore"
                    set frontmost of museScoreProcess to true
                end try
            end try
        end tell
        """,
        """
        tell application "System Events"
            set museScoreProcess to first process whose name contains "MuseScore"
            set frontmost of museScoreProcess to true
        end tell
        """,
    ]

    for script in scripts:
        success, output, error = run_applescript(script)
        if success:
            return True, output, error

    return False, "", "Could not activate MuseScore"


def send_shortcut_macos():
    """Send Cmd+Shift+S shortcut on macOS with multiple fallback methods."""
    scripts = [
        """
        tell application "System Events"
            try
                tell process "mscore"
                    keystroke "s" using {command down, shift down}
                    return true
                end tell
            on error
                return false
            end try
        end tell
        """,
        """
        tell application "System Events"
            try
                tell process "MuseScore 4"
                    keystroke "s" using {command down, shift down}
                    return true
                end tell
            on error
                return false
            end try
        end tell
        """,
        """
        tell application "System Events"
            try
                tell process "MuseScore 3"
                    keystroke "s" using {command down, shift down}
                    return true
                end tell
            on error
                return false
            end try
        end tell
        """,
        """
        tell application "System Events"
            try
                tell process "MuseScore"
                    keystroke "s" using {command down, shift down}
                    return true
                end tell
            on error
                return false
            end try
        end tell
        """,
        """
        tell application "System Events"
            try
                set museScoreProcess to first process whose name is "mscore"
                tell museScoreProcess
                    keystroke "s" using {command down, shift down}
                end tell
                return true
            on error
                return false
            end try
        end tell
        """,
        """
        tell application "System Events"
            try
                set museScoreProcess to first process whose name contains "MuseScore"
                tell museScoreProcess
                    keystroke "s" using {command down, shift down}
                end tell
                return true
            on error
                return false
            end try
        end tell
        """,
        """
        tell application "System Events"
            try
                set museScoreProcess to first process whose name is "mscore"
                tell museScoreProcess
                    keystroke "s" using {command down, shift down}
                end tell
                return true
            on error
                try
                    set museScoreProcess to first process whose name is "MuseScore 4"
                    tell museScoreProcess
                        keystroke "s" using {command down, shift down}
                    end tell
                    return true
                on error
                    try
                        set museScoreProcess to first process whose name contains "MuseScore"
                        tell museScoreProcess
                            keystroke "s" using {command down, shift down}
                        end tell
                        return true
                    on error
                        return false
                    end try
                end try
            end try
        end tell
        """,
    ]

    for script in scripts:
        success, output, error = run_applescript(script)
        if success:
            return True, output, error

    return False, "", "Could not send keyboard shortcut to MuseScore"


class MuseScoreExtractorApp:
    def __init__(self, root, trigger_on_start=False, disable_global_hotkey=False):
        self.root = root
        self.root.title("MuseScore Pitch Extractor")
        self.root.geometry("800x700")

        if EXTRACTION_FUNCTION is None:
            messagebox.showerror(
                "Error",
                "Could not import extraction scripts.\n\n"
                "Please ensure one of these files is in the same directory:\n"
                "- extract_pitches_with_position.py\n"
                "- extract_pitches.py",
            )
            root.destroy()
            return

        self.trigger_on_start = trigger_on_start
        self.disable_global_hotkey = disable_global_hotkey

        self.watch_folder = tk.StringVar()
        self.watching = False
        self.watch_thread = None
        self.processed_files = set()
        self.seen_output_type_files = set()
        self._clear_confirm_queue = queue.Queue()
        self.output_format = tk.StringVar(value="Text")
        self.last_extracted_file = None
        self.delete_previous_var = tk.BooleanVar(value=True)
        self.preferences = self.load_preferences()
        self.visible_programs = list(self.preferences.get("visible_programs", PROGRAM_ORDER))
        self.custom_hotkeys = dict(self.preferences.get("custom_hotkeys", {}))
        self._hotkey_monitor_stop = threading.Event()
        self._last_hotkey_request = 0
        self._pynput_listener = None
        self._last_accepted_watch_event_ts = None
        self._watch_gate_lock = threading.Lock()
        self._ai_flow_lock = threading.Lock()
        self._ai_export_lock = threading.Lock()
        self._ai_export_in_progress = set()

        self.create_widgets()
        self.apply_saved_preferences()
        self.setup_hotkey_request_monitor()
        self.register_global_hotkey()

        if self.trigger_on_start:
            self.root.after(500, self.trigger_save_selection)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def load_preferences(self):
        defaults = {
            "watch_folder": "",
            "watching": False,
            "selected_program": "musescore",
            "visible_programs": list(PROGRAM_ORDER),
            "custom_hotkeys": {},
        }

        if not CONFIG_FILE.exists():
            return defaults

        try:
            raw_content = CONFIG_FILE.read_text(encoding="utf-8")
        except Exception:
            return defaults

        try:
            loaded = json.loads(raw_content)
        except json.JSONDecodeError:
            # Backward compatibility with legacy 2-line format.
            lines = [line.rstrip("\n") for line in raw_content.splitlines()]
            folder = lines[0] if lines else ""
            watching = len(lines) > 1 and lines[1].strip().lower() == "true"
            merged = dict(defaults)
            merged["watch_folder"] = folder
            merged["watching"] = watching
            return merged

        if not isinstance(loaded, dict):
            return defaults

        merged = dict(defaults)
        merged.update({k: v for k, v in loaded.items() if k in merged})

        visible_programs = merged.get("visible_programs")
        if not isinstance(visible_programs, list):
            visible_programs = list(PROGRAM_ORDER)
        visible_programs = [pid for pid in visible_programs if pid in PROGRAM_PROFILES]
        if not visible_programs:
            visible_programs = list(PROGRAM_ORDER)
        merged["visible_programs"] = visible_programs

        selected_program = merged.get("selected_program")
        if selected_program not in PROGRAM_PROFILES:
            selected_program = "musescore"
        if selected_program not in visible_programs:
            selected_program = visible_programs[0]
        merged["selected_program"] = selected_program

        custom_hotkeys = merged.get("custom_hotkeys")
        if not isinstance(custom_hotkeys, dict):
            custom_hotkeys = {}
        normalized_custom = {}
        platform_key = _current_platform_key()
        for pid, raw_hotkey in custom_hotkeys.items():
            if pid not in PROGRAM_PROFILES:
                continue
            normalized, error = _normalize_hotkey_value(raw_hotkey, platform_key)
            if error:
                continue
            normalized_custom[pid] = normalized
        merged["custom_hotkeys"] = normalized_custom
        return merged

    def _get_program_label(self, program_id):
        profile = PROGRAM_PROFILES.get(program_id, PROGRAM_PROFILES["musescore"])
        return profile["label"]

    def _get_selected_program_id(self):
        selected = self.selected_program_var.get() if hasattr(self, "selected_program_var") else ""
        if selected in self.visible_programs:
            return selected
        if self.visible_programs:
            fallback = self.visible_programs[0]
            if hasattr(self, "selected_program_var"):
                self.selected_program_var.set(fallback)
            return fallback
        if hasattr(self, "selected_program_var"):
            self.selected_program_var.set("musescore")
        return "musescore"

    def _get_program_default_hotkey(self, program_id):
        profile = PROGRAM_PROFILES.get(program_id, PROGRAM_PROFILES["musescore"])
        return profile["default_hotkeys"].get(_current_platform_key(), "")

    def _resolve_effective_hotkey(self, program_id=None):
        target_program = program_id or self._get_selected_program_id()
        custom = (self.custom_hotkeys.get(target_program) or "").strip().lower()
        candidate = custom or self._get_program_default_hotkey(target_program)
        if not candidate:
            return None, (
                f"{self._get_program_label(target_program)} has no default shortcut on this platform. "
                "Set one in Settings."
            )
        normalized, error = _normalize_hotkey_value(candidate, _current_platform_key())
        if error:
            return None, error
        return normalized, None

    def save_preferences(self, watching_override=None):
        folder = self.watch_folder.get()
        watching = self.watching if watching_override is None else watching_override
        selected_program = self.selected_program_var.get() if hasattr(self, "selected_program_var") else "musescore"
        if selected_program not in self.visible_programs and self.visible_programs:
            selected_program = self.visible_programs[0]

        self.preferences = {
            "watch_folder": folder,
            "watching": watching,
            "selected_program": selected_program,
            "visible_programs": list(self.visible_programs),
            "custom_hotkeys": dict(self.custom_hotkeys),
        }
        try:
            CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.preferences, f, indent=2)
        except Exception as exc:
            self.log(f"Warning: Could not save preferences: {exc}")

    def apply_saved_preferences(self):
        default_folder = os.path.join(os.path.expanduser("~"), "Documents", "MuseScore4", "Scores")
        saved_folder = self.preferences.get("watch_folder")

        if saved_folder and os.path.exists(saved_folder):
            self.watch_folder.set(saved_folder)
        elif os.path.exists(default_folder):
            self.watch_folder.set(default_folder)
        else:
            self.watch_folder.set(os.path.join(os.path.expanduser("~"), "Documents"))

        selected_program = self.preferences.get("selected_program", "musescore")
        if selected_program not in self.visible_programs and self.visible_programs:
            selected_program = self.visible_programs[0]
        elif selected_program not in PROGRAM_PROFILES:
            selected_program = "musescore"
        self.selected_program_var.set(selected_program)
        self._refresh_program_dropdown()
        self._update_program_dependent_ui()

        if self.preferences.get("watching"):
            folder = self.watch_folder.get().strip()
            if folder and os.path.exists(folder):
                self.root.after(200, self.toggle_watch)

    def create_widgets(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)

        self.selected_program_var = tk.StringVar(value=self.preferences.get("selected_program", "musescore"))
        self.selected_program_display_var = tk.StringVar()
        self.visible_program_vars = {
            pid: tk.BooleanVar(value=pid in self.visible_programs) for pid in PROGRAM_ORDER
        }
        self.custom_hotkey_vars = {
            pid: tk.StringVar(value=self.custom_hotkeys.get(pid, "")) for pid in PROGRAM_ORDER
        }

        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.main_tab = ttk.Frame(notebook, padding="10")
        self.settings_tab = ttk.Frame(notebook, padding="10")
        notebook.add(self.main_tab, text="Main")
        notebook.add(self.settings_tab, text="Settings")

        self._build_main_tab()
        self._build_settings_tab()

        self.selected_program_dropdown.bind("<<ComboboxSelected>>", self._on_selected_program_changed)
        self._refresh_program_dropdown()
        self._update_program_dependent_ui()

    def _build_main_tab(self):
        main_tab = self.main_tab
        main_tab.columnconfigure(0, weight=1)
        main_tab.rowconfigure(2, weight=1)

        title_label = ttk.Label(
            main_tab,
            text="Music Clipboard Extractor",
            font=("Arial", 16, "bold"),
        )
        title_label.grid(row=0, column=0, pady=(0, 20), sticky=tk.W)

        watch_frame = ttk.LabelFrame(main_tab, text="Auto-Process Saved Selections", padding="10")
        watch_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        watch_frame.columnconfigure(1, weight=1)

        ttk.Label(watch_frame, text="Watch Folder:").grid(row=0, column=0, sticky=tk.W, padx=5)
        watch_entry = ttk.Entry(watch_frame, textvariable=self.watch_folder, width=50, state="readonly")
        watch_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(watch_frame, text="Browse...", command=self.browse_watch_folder).grid(row=0, column=2, padx=5)

        format_frame = ttk.Frame(watch_frame)
        format_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(10, 0))
        ttk.Label(format_frame, text="Output Format:").grid(row=0, column=0, padx=5)
        format_dropdown = ttk.Combobox(
            format_frame,
            textvariable=self.output_format,
            values=["Text", "MIDI"],
            state="readonly",
        )
        format_dropdown.current(0)
        format_dropdown.grid(row=0, column=1, padx=5)

        program_frame = ttk.Frame(watch_frame)
        program_frame.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(8, 0))
        ttk.Label(program_frame, text="Target Program:").grid(row=0, column=0, padx=5)
        self.selected_program_dropdown = ttk.Combobox(
            program_frame,
            textvariable=self.selected_program_display_var,
            state="readonly",
            width=22,
        )
        self.selected_program_dropdown.grid(row=0, column=1, padx=5)

        automation_frame = ttk.Frame(watch_frame)
        automation_frame.grid(row=3, column=0, columnspan=3, pady=5, sticky=tk.W)

        self.save_selection_button = ttk.Button(
            automation_frame,
            text="Trigger Save/Export",
            command=self.trigger_save_selection,
        )
        self.save_selection_button.grid(row=0, column=0, padx=5)

        hotkey_available = (IS_MACOS and PYNPUT_AVAILABLE) or (not IS_MACOS and KEYBOARD_AVAILABLE)
        if hotkey_available:
            hotkey_label = ttk.Label(
                automation_frame,
                text=(
                    f"Global Hotkey: {_format_hotkey_label(GLOBAL_HOTKEY)} "
                    "(background listener keeps it active even when the GUI is closed)"
                ),
                foreground="green",
                font=("Arial", 9, "bold"),
            )
            hotkey_label.grid(row=0, column=1, padx=10)
        else:
            install_hint = "pip install pynput" if IS_MACOS else "pip install keyboard"
            hotkey_label = ttk.Label(
                automation_frame,
                text=f"(Install '{'pynput' if IS_MACOS else 'keyboard'}' for global hotkey: {install_hint})",
                foreground="gray",
                font=("Arial", 8),
            )
            hotkey_label.grid(row=0, column=1, padx=5)

        automation_ready = IS_MACOS or (IS_WINDOWS and PYAUTOGUI_AVAILABLE and PYWINAUTO_AVAILABLE)
        if not automation_ready:
            self.save_selection_button.config(state="disabled")
            missing_libs = []
            if IS_WINDOWS:
                if not PYWINAUTO_AVAILABLE:
                    missing_libs.append("pywinauto")
                if not PYAUTOGUI_AVAILABLE:
                    missing_libs.append("pyautogui")
                if missing_libs:
                    ttk.Label(
                        automation_frame,
                        text=f"(Install: pip install {' '.join(missing_libs)})",
                        foreground="gray",
                        font=("Arial", 8),
                    ).grid(row=1, column=0, columnspan=2, padx=5, sticky=tk.W)
            elif not IS_MACOS:
                ttk.Label(
                    automation_frame,
                    text="(Automation is supported on macOS and Windows)",
                    foreground="gray",
                    font=("Arial", 8),
                ).grid(row=1, column=0, columnspan=2, padx=5, sticky=tk.W)

        self.watch_button = ttk.Button(watch_frame, text="Start Watching", command=self.toggle_watch)
        self.watch_button.grid(row=4, column=0, columnspan=3, pady=10)

        self.watch_status_label = ttk.Label(watch_frame, text="Status: Not watching", foreground="gray")
        self.watch_status_label.grid(row=5, column=0, columnspan=3)

        self.instructions_label = ttk.Label(watch_frame, justify=tk.LEFT, foreground="gray")
        self.instructions_label.grid(row=6, column=0, columnspan=3, pady=10, sticky=tk.W)

        output_frame = ttk.LabelFrame(main_tab, text="Output", padding="10")
        output_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)

        self.output_text = scrolledtext.ScrolledText(output_frame, height=15, width=80, wrap=tk.WORD)
        self.output_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        button_frame = ttk.Frame(output_frame)
        button_frame.grid(row=1, column=0, pady=5)

        ttk.Button(button_frame, text="Clear Output", command=self.clear_output).grid(row=0, column=0, padx=5)

        self.open_location_button = ttk.Button(
            button_frame,
            text="Open File Location",
            command=self.open_file_location,
            state="disabled",
        )
        self.open_location_button.grid(row=0, column=1, padx=5)

        ttk.Checkbutton(
            button_frame,
            text="Auto-delete previous extraction",
            variable=self.delete_previous_var,
        ).grid(row=0, column=2, padx=5)

    def _build_settings_tab(self):
        settings_tab = self.settings_tab
        settings_tab.columnconfigure(0, weight=1)

        settings_frame = ttk.LabelFrame(settings_tab, text="Program Visibility & Hotkeys", padding="10")
        settings_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N), pady=5)
        settings_frame.columnconfigure(2, weight=1)

        ttk.Label(
            settings_frame,
            text=(
                "Control which programs appear in the main dropdown and override the save/export shortcut per program.\n"
                "Leave custom hotkey blank to use platform default."
            ),
            justify=tk.LEFT,
            foreground="gray",
        ).grid(row=0, column=0, columnspan=5, sticky=tk.W, pady=(0, 12))

        ttk.Label(settings_frame, text="Program", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky=tk.W, padx=5)
        ttk.Label(settings_frame, text="Show", font=("Arial", 10, "bold")).grid(row=1, column=1, sticky=tk.W, padx=5)
        ttk.Label(settings_frame, text="Custom Hotkey", font=("Arial", 10, "bold")).grid(row=1, column=2, sticky=tk.W, padx=5)
        ttk.Label(settings_frame, text="Default", font=("Arial", 10, "bold")).grid(row=1, column=3, sticky=tk.W, padx=5)
        ttk.Label(settings_frame, text="Action", font=("Arial", 10, "bold")).grid(row=1, column=4, sticky=tk.W, padx=5)

        row = 2
        for pid in PROGRAM_ORDER:
            ttk.Label(settings_frame, text=self._get_program_label(pid)).grid(row=row, column=0, sticky=tk.W, padx=5, pady=4)
            ttk.Checkbutton(settings_frame, variable=self.visible_program_vars[pid]).grid(
                row=row, column=1, sticky=tk.W, padx=5, pady=4
            )
            ttk.Entry(settings_frame, textvariable=self.custom_hotkey_vars[pid], width=24).grid(
                row=row, column=2, sticky=(tk.W, tk.E), padx=5, pady=4
            )
            default_hotkey = self._get_program_default_hotkey(pid)
            default_label = _format_hotkey_label(default_hotkey) if default_hotkey else "No default"
            ttk.Label(settings_frame, text=default_label, foreground="gray").grid(
                row=row, column=3, sticky=tk.W, padx=5, pady=4
            )
            ttk.Button(
                settings_frame,
                text="Reset",
                command=lambda target=pid: self._reset_custom_hotkey(target),
            ).grid(row=row, column=4, sticky=tk.W, padx=5, pady=4)
            row += 1

        ttk.Button(settings_tab, text="Save Settings", command=self._save_settings).grid(
            row=1, column=0, sticky=tk.E, pady=8
        )

    def _refresh_program_dropdown(self):
        available_ids = [pid for pid in PROGRAM_ORDER if pid in self.visible_programs]
        if not available_ids:
            available_ids = ["musescore"]
        labels = [self._get_program_label(pid) for pid in available_ids]
        self.selected_program_dropdown["values"] = labels

        selected_id = self.selected_program_var.get()
        if selected_id not in available_ids:
            selected_id = available_ids[0]
            self.selected_program_var.set(selected_id)
        self.selected_program_display_var.set(self._get_program_label(selected_id))

    def _on_selected_program_changed(self, _event=None):
        selected_label = self.selected_program_display_var.get()
        program_id = next(
            (pid for pid in PROGRAM_ORDER if self._get_program_label(pid) == selected_label),
            "musescore",
        )
        if program_id not in self.visible_programs and self.visible_programs:
            program_id = self.visible_programs[0]
        self.selected_program_var.set(program_id)
        self._update_program_dependent_ui()
        self.save_preferences()

    def _build_instruction_text(self):
        selected_program = self._get_selected_program_id()
        selected_label = self._get_program_label(selected_program)
        effective_hotkey, hotkey_error = self._resolve_effective_hotkey(selected_program)
        shortcut_hint = _format_hotkey_label(effective_hotkey) if effective_hotkey else f"Not configured ({hotkey_error})"
        return (
            "Instructions:\n"
            "1. Auto Mode:\n"
            "   - Set the watch folder (where scores are saved), click 'Start Watching'\n"
            "   - In your notation/DAW app: select content to export/save\n"
            f"   - Click 'Trigger Save/Export in {selected_label}' (or use the app's own save/export path)\n"
            "   - Save in the watch folder\n"
            f"   - Shortcut reminder for {selected_label}: {shortcut_hint}\n"
            "\n"
            "Note: Watched-file AI opening remains MuseScore-only."
        )

    def _update_program_dependent_ui(self):
        selected_program = self._get_selected_program_id()
        selected_label = self._get_program_label(selected_program)
        self.save_selection_button.config(text=f"Trigger Save/Export in {selected_label}")
        self.instructions_label.config(text=self._build_instruction_text())

    def _reset_custom_hotkey(self, program_id):
        self.custom_hotkey_vars[program_id].set("")

    def _save_settings(self):
        visible = [pid for pid in PROGRAM_ORDER if self.visible_program_vars[pid].get()]
        if not visible:
            for pid in PROGRAM_ORDER:
                self.visible_program_vars[pid].set(pid in self.visible_programs)
            messagebox.showerror("Invalid Settings", "At least one program must remain visible.")
            return

        platform_key = _current_platform_key()
        normalized_custom = {}
        for pid in PROGRAM_ORDER:
            raw_hotkey = (self.custom_hotkey_vars[pid].get() or "").strip()
            if not raw_hotkey:
                normalized_custom[pid] = ""
                continue
            normalized, error = _normalize_hotkey_value(raw_hotkey, platform_key)
            if error:
                messagebox.showerror(
                    "Invalid Hotkey",
                    f"{self._get_program_label(pid)} hotkey is invalid:\n{error}",
                )
                return
            normalized_custom[pid] = normalized

        self.visible_programs = visible
        self.custom_hotkeys = normalized_custom

        selected_program = self._get_selected_program_id()
        if selected_program not in self.visible_programs:
            self.selected_program_var.set(self.visible_programs[0])

        for pid in PROGRAM_ORDER:
            self.custom_hotkey_vars[pid].set(self.custom_hotkeys.get(pid, ""))

        self._refresh_program_dropdown()
        self._update_program_dependent_ui()
        self.save_preferences()
        self.log("Settings saved.")

    def browse_watch_folder(self):
        folder = filedialog.askdirectory(title="Select Folder to Watch")
        if folder:
            self.watch_folder.set(folder)
            self.save_preferences()

    def log(self, message):
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)
        self.root.update_idletasks()

    def clear_output(self):
        self.output_text.delete(1.0, tk.END)

    def _reveal_file_in_folder(self, file_path):
        """Reveal a file in the system file manager (Finder/Explorer) without changing app state."""
        if not file_path or not os.path.exists(file_path):
            return
        try:
            if IS_MACOS:
                subprocess.run(["open", "-R", file_path])
            elif IS_WINDOWS:
                subprocess.run(["explorer", "/select,", os.path.normpath(file_path)])
            else:
                subprocess.run(["xdg-open", os.path.dirname(file_path)])
        except Exception:
            try:
                folder = os.path.dirname(file_path)
                if IS_WINDOWS:
                    os.startfile(folder)
                else:
                    subprocess.run(["open" if IS_MACOS else "xdg-open", folder])
            except Exception:
                pass

    def _bring_app_to_front(self):
        """Raise and focus the app window."""
        def do_bring():
            try:
                self.root.attributes("-topmost", True)
                self.root.lift()
                self.root.focus_force()
                self.root.attributes("-topmost", False)
            except Exception:
                pass

        self.root.after(0, do_bring)

    def _process_clear_confirm_queue(self):
        """Run on main thread: show warning and get user OK/Cancel when output folder has wrong-extension files."""
        try:
            request = self._clear_confirm_queue.get_nowait()
        except queue.Empty:
            return
        dest_dir, wrong_ext_files, _all_files, _src_path, output_ext, response_queue = request
        format_name = "MIDI" if output_ext == ".mid" else "Text"
        expected_ext = ".mid" if output_ext == ".mid" else ".txt"
        file_list = "\n".join(f"  • {os.path.basename(p)}" for p in wrong_ext_files)
        msg = (
            f"The output folder contains file(s) that are not {format_name} format (expected {expected_ext}):\n\n"
            f"{file_list}\n\n"
            "Everything in the output folder will be deleted, then the new file will be moved in.\n\n"
            "Continue?"
        )
        ok = messagebox.askokcancel("Clear output folder?", msg)
        try:
            response_queue.put(ok)
        except Exception:
            pass

    def _clear_output_folder_and_move(self, src_path, output_ext):
        """
        Delete everything in the output folder for this format, then move src_path into it.
        If any existing file has an extension other than the selected output format, show a warning and Cancel option.
        Returns the destination path on success, None on cancel or failure.
        """
        if not src_path or not os.path.exists(src_path):
            return None
        dest_dir = Path(MIDI_OUTPUT_DIR) if output_ext == ".mid" else Path(OUTPUT_DIR)
        try:
            dest_dir.mkdir(parents=True, exist_ok=True)
            all_entries = list(dest_dir.iterdir())
            all_files = [p for p in all_entries if p.is_file()]
            wrong_ext_files = [p for p in all_files if p.suffix.lower() != output_ext]

            if wrong_ext_files:
                response_queue = queue.Queue()
                self._clear_confirm_queue.put(
                    (dest_dir, wrong_ext_files, all_files, src_path, output_ext, response_queue)
                )
                self.root.after(0, self._process_clear_confirm_queue)
                while self.watching:
                    try:
                        proceed = response_queue.get(timeout=0.5)
                        break
                    except queue.Empty:
                        continue
                else:
                    return None
                if not proceed:
                    return None

            for p in all_files:
                try:
                    os.remove(p)
                except OSError:
                    pass

            dest_path = dest_dir / os.path.basename(src_path)
            shutil.move(src_path, dest_path)
            return str(dest_path)
        except Exception:
            return None

    def open_file_location(self):
        if self.last_extracted_file and os.path.exists(self.last_extracted_file):
            try:
                if IS_MACOS:
                    subprocess.run(["open", "-R", self.last_extracted_file])
                elif IS_WINDOWS:
                    subprocess.run(["explorer", "/select,", os.path.normpath(self.last_extracted_file)])
                else:
                    subprocess.run(["xdg-open", os.path.dirname(self.last_extracted_file)])
            except Exception:
                try:
                    folder = os.path.dirname(self.last_extracted_file)
                    if IS_WINDOWS:
                        os.startfile(folder)
                    else:
                        subprocess.run(["open" if IS_MACOS else "xdg-open", folder])
                except Exception as e2:
                    self.log(f"Error opening file location: {str(e2)}\n")
                    messagebox.showerror("Error", f"Could not open file location:\n{str(e2)}")
        else:
            messagebox.showwarning("Warning", "No extracted file location available.")

    def _delete_previous_extracted_file(self, new_path):
        previous_path = self.last_extracted_file
        if not previous_path or previous_path == new_path:
            return
        try:
            if os.path.exists(previous_path):
                os.remove(previous_path)
                self.log(f"Deleted previous extraction: {previous_path}")
        except Exception as exc:
            self.log(f"Failed to delete previous extraction ({previous_path}): {exc}")

    def _handle_successful_extraction(self, extracted_path):
        if not extracted_path:
            return
        if self.delete_previous_var.get():
            self._delete_previous_extracted_file(extracted_path)
        self.last_extracted_file = extracted_path
        self.root.after(0, lambda: self.open_location_button.config(state="normal"))
        self.root.after(0, self.open_file_location)

    def _show_error_async(self, title, message):
        self.root.after(0, lambda: messagebox.showerror(title, message))

    def _should_accept_new_file(self, now_monotonic):
        with self._watch_gate_lock:
            last = self._last_accepted_watch_event_ts
            if last is not None and now_monotonic - last < 5.0:
                return False
            self._last_accepted_watch_event_ts = now_monotonic
            return True

    def _is_claude_running_macos(self):
        if not IS_MACOS:
            return False

        script = """
        tell application "System Events"
            try
                set _ to first process whose name is "Claude"
                return true
            on error
                return false
            end try
        end tell
        """
        success, output, _ = run_applescript(script)
        if success and output.strip().lower() == "true":
            return True

        if PSUTIL_AVAILABLE:
            try:
                for proc in psutil.process_iter(["name"]):
                    try:
                        name = (proc.info.get("name") or "").lower()
                        if name == "claude" or "claude" in name:
                            return True
                    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                        continue
            except Exception:
                pass
        return False

    def _open_file_in_musescore(self, file_path):
        if not IS_MACOS:
            return False, "MuseScore open automation is only supported on macOS for this flow."
        try:
            result = subprocess.run(
                ["open", "-a", "MuseScore 4", file_path],
                capture_output=True,
                text=True,
            )
            if result.returncode != 0:
                err = result.stderr.strip() or result.stdout.strip() or "Unknown error opening MuseScore 4."
                return False, err
            return True, ""
        except Exception as exc:
            return False, str(exc)

    def _start_musescore_websocket_plugin(self):
        if not IS_MACOS:
            return False, "MuseScore plugin shortcut automation is only supported on macOS."

        hotkey, hotkey_error = _normalize_hotkey_value("ctrl+cmd+p", "mac")
        if not hotkey:
            return False, hotkey_error or "Invalid MuseScore plugin shortcut."

        last_error = "MuseScore process not found."
        for _ in range(20):
            found, _, find_error = self._find_program_window_macos("musescore")
            if found:
                activated, _, activate_error = self._activate_program_window_macos("musescore")
                if activated:
                    time.sleep(0.15)
                    sent, _, send_error = self._send_hotkey_macos(hotkey)
                    if sent:
                        return True, ""
                    last_error = send_error or "Could not send macOS shortcut"
                else:
                    last_error = activate_error or "Could not activate MuseScore."
            else:
                last_error = find_error or "MuseScore process not found."
            time.sleep(0.2)

        return False, last_error

    def _send_prompt_to_claude(self, prompt_text, export_midi_path=None, source_score_path=None):
        if not IS_MACOS:
            return False, "Claude automation is only supported on macOS for this flow."

        user_prompt = prompt_text.strip()
        if export_midi_path and source_score_path:
            user_prompt = (
                f"{user_prompt}\n\n"
                "When you are completely finished with all edits, as the final step run a local MuseScore CLI export "
                "(do not do this early). Use execute_bash to run exactly one of these commands:\n"
                f'"/Applications/MuseScore 4.app/Contents/MacOS/mscore" -o "{export_midi_path}" "{source_score_path}"\n'
                f'or\n"/Applications/MuseScore 4.app/Contents/MacOS/MuseScore4" -o "{export_midi_path}" "{source_score_path}"\n'
                "Overwrite the output file if it already exists."
            )

        full_prompt = f"Connect to musescore and {user_prompt}"

        try:
            subprocess.run(["pbcopy"], input=full_prompt, text=True, check=True)
        except Exception as exc:
            return False, f"Failed to set clipboard text: {exc}"

        script = """
        tell application "System Events"
            try
                set claudeProcess to first process whose name is "Claude"
            on error
                error "Claude is not running."
            end try
            set frontmost of claudeProcess to true
            delay 0.2
            keystroke "v" using {command down}
            delay 0.1
            key code 36
        end tell
        """
        success, _, error = run_applescript(script)
        if not success:
            return False, error or "Failed to send prompt to Claude."

        return True, ""

    def _auto_export_ai_result_to_midi_thread(self, file_path, export_midi_path, baseline_mtime, prompt_sent_time=None):
        try:
            if prompt_sent_time is None:
                prompt_sent_time = time.time()

            deadline = prompt_sent_time + 1800.0
            stable_since = None
            last_state = None
            threshold_mtime = max(float(baseline_mtime or 0), float(prompt_sent_time))

            self.log(
                "Waiting for Claude final export via MCP tool "
                f"(up to 30 minutes): {os.path.basename(export_midi_path)}"
            )

            while time.time() < deadline:
                if not os.path.exists(export_midi_path):
                    time.sleep(1.0)
                    continue

                try:
                    mtime = os.path.getmtime(export_midi_path)
                    size = os.path.getsize(export_midi_path)
                except OSError:
                    time.sleep(1.0)
                    continue

                if mtime < threshold_mtime or size <= 0:
                    time.sleep(1.0)
                    continue

                state = (mtime, size)
                if state != last_state:
                    last_state = state
                    stable_since = time.time()
                elif stable_since is not None and time.time() - stable_since >= 2.0:
                    self.log(f"Detected Claude-exported MIDI: {export_midi_path}")
                    self._handle_successful_extraction(export_midi_path)
                    return

                time.sleep(1.0)

            self.log(
                "AI export timed out waiting for Claude final export. "
                f"Expected file: {export_midi_path}"
            )
        finally:
            with self._ai_export_lock:
                self._ai_export_in_progress.discard(file_path)

    def _start_auto_export_ai_result_to_midi(self, file_path, export_midi_path, baseline_mtime, prompt_sent_time=None):
        with self._ai_export_lock:
            if file_path in self._ai_export_in_progress:
                self.log(f"AI auto-export already in progress for: {os.path.basename(file_path)}")
                return
            self._ai_export_in_progress.add(file_path)

        thread = threading.Thread(
            target=self._auto_export_ai_result_to_midi_thread,
            args=(file_path, export_midi_path, baseline_mtime, prompt_sent_time),
            daemon=True,
        )
        thread.start()

    def _run_ai_edit_flow(self, file_path, prompt_text):
        with self._ai_flow_lock:
            if not IS_MACOS:
                return

            if not self._is_claude_running_macos():
                error_msg = (
                    "Claude must already be running before this flow can open MuseScore 4.\n\n"
                    "Please open Claude Desktop and try with the next detected file."
                )
                self.log(f"Error: {error_msg}")
                self._show_error_async("Claude Not Running", error_msg)
                return

            opened, open_error = self._open_file_in_musescore(file_path)
            if not opened:
                error_msg = (
                    f"Failed to open MuseScore 4 for file:\n{file_path}\n\n"
                    f"Details: {open_error}\n\nClaude send was canceled."
                )
                self.log(f"Error: {error_msg}")
                self._show_error_async("MuseScore Open Failed", error_msg)
                return

            time.sleep(0.3)
            plugin_started, plugin_error = self._start_musescore_websocket_plugin()
            if not plugin_started:
                error_msg = (
                    "Opened MuseScore 4, but failed to start the websocket plugin shortcut (Ctrl+Cmd+P).\n\n"
                    f"Details: {plugin_error}\n\nClaude send was canceled."
                )
                self.log(f"Error: {error_msg}")
                self._show_error_async("MuseScore Plugin Start Failed", error_msg)
                return

            os.makedirs(MIDI_OUTPUT_DIR, exist_ok=True)
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            export_midi_path = os.path.join(MIDI_OUTPUT_DIR, f"{base_name}_ai.mid")
            baseline_mtime = 0.0
            if os.path.exists(export_midi_path):
                try:
                    baseline_mtime = os.path.getmtime(export_midi_path)
                except OSError:
                    baseline_mtime = 0.0

            time.sleep(0.1)
            sent, send_error = self._send_prompt_to_claude(
                prompt_text,
                export_midi_path=export_midi_path,
                source_score_path=file_path,
            )
            if not sent:
                error_msg = f"Failed to send prompt to Claude: {send_error}"
                self.log(f"Error: {error_msg}")
                self._show_error_async("Claude Send Failed", error_msg)
                return

            self.log(f"Sent AI prompt to Claude for: {os.path.basename(file_path)}")
            self._start_auto_export_ai_result_to_midi(
                file_path,
                export_midi_path=export_midi_path,
                baseline_mtime=baseline_mtime,
                prompt_sent_time=time.time(),
            )

    def handle_new_score_file(self, file_path):
        if not file_path or not os.path.exists(file_path):
            return

        extension = Path(file_path).suffix.lower()
        self.log(f"Detected new file: {os.path.basename(file_path)}")

        if extension in EXTRACTABLE_SCORE_EXTENSIONS:
            self.extract_file(file_path)
        elif extension in (".mid", ".midi"):
            self.log(f"Skipping extraction for MIDI input: {os.path.basename(file_path)}")

        if not IS_MACOS:
            return

        prompt = simpledialog.askstring(
            "AI Edit Prompt",
            (
                f"Enter the AI edit prompt for:\n{os.path.basename(file_path)}\n\n"
                "The app will send:\n"
                "Connect to musescore and <your prompt>"
            ),
            parent=self.root,
        )

        if prompt is None or not prompt.strip():
            self.log(f"Skipped Claude send for {os.path.basename(file_path)} (empty/canceled prompt).")
            return

        thread = threading.Thread(
            target=self._run_ai_edit_flow,
            args=(file_path, prompt.strip()),
            daemon=True,
        )
        thread.start()

    def extract_file(self, file_path):
        if not file_path:
            messagebox.showwarning("Warning", "No file provided for auto-processing.")
            return

        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"File not found: {file_path}")
            return

        thread = threading.Thread(target=self._extract_thread, args=(file_path,), daemon=True)
        thread.start()

    def _extract_thread(self, file_path):
        output_format = self.output_format.get().strip().lower()
 
        self.log(f"\n{'=' * 60}")
        self.log(f"Processing: {os.path.basename(file_path)}")
        self.log(f"Output format: {output_format.upper()}")
        self.log(f"{'=' * 60}\n")

        try:
            if output_format == "midi":
                if MIDI_EXTRACTION_FUNCTION is None:
                    error_msg = (
                        "MIDI extraction function not available. Please ensure extract_midi.py is in the same directory."
                    )
                    self.log(f"{error_msg}\n")
                    self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
                    return

                os.makedirs(MIDI_OUTPUT_DIR, exist_ok=True)

                base_name = os.path.splitext(os.path.basename(file_path))[0]
                output_file = os.path.join(MIDI_OUTPUT_DIR, base_name + ".mid")

                try:
                    midi_path = MIDI_EXTRACTION_FUNCTION(file_path, output_file)

                    if midi_path and os.path.exists(midi_path):
                        self.log("Successfully extracted MIDI!")
                        self.log(f"MIDI file saved to: {midi_path}\n")
                        self._handle_successful_extraction(midi_path)
                    else:
                        error_msg = "Failed to extract MIDI file."
                        self.log(f"{error_msg}\n")
                        self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
                except Exception as e:
                    error_msg = f"Error extracting MIDI: {str(e)}"
                    self.log(f"{error_msg}\n")
                    import traceback

                    self.log(traceback.format_exc())
                    self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
            else:
                os.makedirs(OUTPUT_DIR, exist_ok=True)

                base_name = os.path.splitext(os.path.basename(file_path))[0]
                if EXTRACTION_SCRIPT == "extract_pitches_with_position":
                    output_file = os.path.join(OUTPUT_DIR, base_name + "_pitches_with_position.txt")
                else:
                    output_file = os.path.join(OUTPUT_DIR, base_name + "_pitches.txt")

                result = EXTRACTION_FUNCTION(file_path, output_file, debug=False)

                if isinstance(result, tuple) and len(result) == 2:
                    notes, actual_output_path = result
                    output_file = actual_output_path
                else:
                    notes = result

                if notes:
                    self.log(f"Successfully extracted {len(notes)} notes!")
                    self.log(f"Output saved to: {output_file}\n")
                    self._handle_successful_extraction(output_file)

                    self.log("First 10 notes:")
                    for i, (pitch, position, tick) in enumerate(notes[:10], 1):
                        if tick is not None:
                            self.log(f"  {i}. {pitch} | {position} | (tick: {tick})")
                        else:
                            self.log(f"  {i}. {pitch} | {position}")

                    if len(notes) > 10:
                        self.log(f"  ... and {len(notes) - 10} more\n")
                else:
                    self.log("No notes extracted. Please check the file format.\n")
                    self.root.after(0, lambda: messagebox.showerror("Error", "No notes were extracted from the file."))
        except Exception as e:
            error_msg = f"Error processing file: {str(e)}"
            self.log(f"{error_msg}\n")
            import traceback

            self.log(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))

    def toggle_watch(self):
        if not self.watching:
            folder = self.watch_folder.get().strip()
            if not folder or not os.path.exists(folder):
                messagebox.showerror("Error", "Please select a valid folder to watch.")
                return

            self.watching = True
            self.watch_button.config(text="Stop Watching")
            self.watch_status_label.config(
                text=f"Status: Watching '{os.path.basename(folder)}'", foreground="green"
            )
            self.log(f"Started watching folder: {folder}\n")
            self.save_preferences()

            self.watch_thread = threading.Thread(target=self._watch_folder, args=(folder,), daemon=True)
            self.watch_thread.start()
        else:
            self.watching = False
            self.watch_button.config(text="Start Watching")
            self.watch_status_label.config(text="Status: Not watching", foreground="gray")
            self.log("Stopped watching folder.\n")
            self.save_preferences()

    def _watch_folder(self, folder):
        initial_files = set()
        for file in os.listdir(folder):
            if file.lower().endswith(WATCHED_SCORE_EXTENSIONS):
                full_path = os.path.join(folder, file)
                initial_files.add(full_path)

        self.processed_files.update(initial_files)

        output_ext = None
        try:
            fmt = (self.output_format.get() or "").strip().lower()
            output_ext = ".mid" if fmt == "midi" else ".txt"
        except Exception:
            output_ext = ".txt"

        initial_output_files = set()
        for file in os.listdir(folder):
            if file.endswith(output_ext):
                if file.lower().endswith(WATCHED_SCORE_EXTENSIONS):
                    continue
                full_path = os.path.join(folder, file)
                initial_output_files.add(full_path)
        self.seen_output_type_files.update(initial_output_files)

        while self.watching:
            try:
                current_files = set()
                for file in os.listdir(folder):
                    if file.lower().endswith(WATCHED_SCORE_EXTENSIONS):
                        full_path = os.path.join(folder, file)
                        current_files.add(full_path)

                        if full_path not in self.processed_files:
                            time.sleep(0.5)

                            try:
                                mod_time = os.path.getmtime(full_path)
                                if time.time() - mod_time > 1:
                                    self.processed_files.add(full_path)
                                    if self._should_accept_new_file(time.monotonic()):
                                        self.root.after(0, lambda f=full_path: self.handle_new_score_file(f))
                            except OSError:
                                pass

                self.processed_files.intersection_update(current_files)

                try:
                    fmt = (self.output_format.get() or "").strip().lower()
                    output_ext = ".mid" if fmt == "midi" else ".txt"
                except Exception:
                    output_ext = ".txt"

                current_output_files = set()
                for file in os.listdir(folder):
                    if file.endswith(output_ext):
                        if file.lower().endswith(WATCHED_SCORE_EXTENSIONS):
                            continue
                        full_path = os.path.join(folder, file)
                        current_output_files.add(full_path)
                        if full_path not in self.seen_output_type_files:
                            dest_path = self._clear_output_folder_and_move(full_path, output_ext)
                            if dest_path is not None:
                                self.seen_output_type_files.add(full_path)
                                self.root.after(0, self._bring_app_to_front)
                                self.root.after(0, lambda p=dest_path: self._handle_successful_extraction(p))
                                self.log(f"Cleared output folder and moved {os.path.basename(full_path)} to: {dest_path}")
                            else:
                                self.log(f"Skipped or failed moving {os.path.basename(full_path)} to output folder")

                self.seen_output_type_files.intersection_update(current_output_files)
                time.sleep(1)

            except Exception as e:
                if self.watching:
                    self.log(f"Error watching folder: {str(e)}\n")
                time.sleep(2)

    def _find_program_window_macos(self, program_id):
        profile = PROGRAM_PROFILES[program_id]

        for process_name in profile["mac_process_names"]:
            script = f"""
            tell application "System Events"
                try
                    set targetProcess to first process whose name is "{process_name}"
                    return name of targetProcess
                on error
                    return ""
                end try
            end tell
            """
            success, output, _ = run_applescript(script)
            if success and output and output.strip():
                return True, output.strip(), ""

        if PSUTIL_AVAILABLE:
            try:
                for proc in psutil.process_iter(["name"]):
                    try:
                        proc_name = proc.info.get("name") or ""
                        proc_lower = proc_name.lower()
                        if any(keyword in proc_lower for keyword in profile["mac_contains"]):
                            return True, proc_name, ""
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        continue
            except Exception:
                pass

        return False, "", f"{profile['label']} process not found"

    def _activate_program_window_macos(self, program_id):
        profile = PROGRAM_PROFILES[program_id]

        for app_name in profile["mac_app_names"]:
            script = f"""
            tell application "{app_name}"
                activate
            end tell
            """
            success, output, error = run_applescript(script)
            if success:
                return True, output, error

        for process_name in profile["mac_process_names"]:
            script = f"""
            tell application "System Events"
                try
                    set targetProcess to first process whose name is "{process_name}"
                    set frontmost of targetProcess to true
                    return true
                on error
                    return false
                end try
            end tell
            """
            success, output, error = run_applescript(script)
            if success and output.strip().lower() == "true":
                return True, output, error

        return False, "", f"Could not activate {profile['label']}"

    def _send_hotkey_macos(self, normalized_hotkey):
        modifiers, key = _split_normalized_hotkey(normalized_hotkey)
        key_to_send = " " if key == "space" else key
        modifier_tokens = []
        for mod in modifiers:
            if mod == "cmd":
                modifier_tokens.append("command down")
            elif mod == "ctrl":
                modifier_tokens.append("control down")
            elif mod == "alt":
                modifier_tokens.append("option down")
            elif mod == "shift":
                modifier_tokens.append("shift down")

        if modifier_tokens:
            using_clause = " using {" + ", ".join(modifier_tokens) + "}"
        else:
            using_clause = ""

        script = f"""
        tell application "System Events"
            keystroke "{key_to_send}"{using_clause}
            return true
        end tell
        """
        success, output, error = run_applescript(script)
        if success and output.strip().lower() == "true":
            return True, output, error
        return False, output, error or "Could not send macOS shortcut"

    def _find_program_window_windows(self, program_id):
        profile = PROGRAM_PROFILES[program_id]
        app = None
        methods_tried = []

        if program_id == "musescore":
            for backend in ["uia", "win32", None]:
                try:
                    if backend:
                        app = Application(backend=backend).connect(path="MuseScore4.exe")
                        methods_tried.append(f"{backend} by process")
                    else:
                        app = Application().connect(path="MuseScore4.exe")
                        methods_tried.append("default by process")
                    return app, methods_tried
                except Exception as e:
                    methods_tried.append(f"{backend or 'default'} process failed: {str(e)[:50]}")

        title_keywords = profile["windows_title_keywords"]
        if title_keywords:
            pattern = title_keywords[0]
            try:
                app = Application(backend="uia").connect(title_re=f".*{pattern}.*")
                methods_tried.append("UIA by title regex")
                return app, methods_tried
            except Exception as e:
                methods_tried.append(f"title regex failed: {str(e)[:50]}")

        if PSUTIL_AVAILABLE:
            try:
                for proc in psutil.process_iter(["pid", "name", "exe"]):
                    try:
                        proc_name = (proc.info.get("name") or "").lower()
                        proc_exe = (proc.info.get("exe") or "").lower()
                        if any(keyword in proc_name or keyword in proc_exe for keyword in profile["windows_process_keywords"]):
                            app = Application(backend="uia").connect(process=proc.info["pid"])
                            methods_tried.append(f"process enumeration: {proc.info.get('name')}")
                            return app, methods_tried
                    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                        continue
            except Exception as e:
                methods_tried.append(f"process search failed: {str(e)[:50]}")

        if WIN32_AVAILABLE:
            try:
                def enum_handler(hwnd, ctx):
                    if not win32gui.IsWindowVisible(hwnd):
                        return
                    window_text = (win32gui.GetWindowText(hwnd) or "").lower()
                    if any(keyword in window_text for keyword in title_keywords):
                        ctx.append(hwnd)

                handles = []
                win32gui.EnumWindows(enum_handler, handles)
                if handles:
                    app = Application().connect(handle=handles[0])
                    methods_tried.append("Windows API by title")
                    return app, methods_tried
            except Exception as e:
                methods_tried.append(f"Windows API failed: {str(e)[:50]}")

        return None, methods_tried

    def _send_hotkey_windows(self, main_window, normalized_hotkey):
        shortcut_sent = False
        last_error = None
        pywinauto_sequence = _hotkey_to_windows_pywinauto(normalized_hotkey)
        modifiers, key = _split_normalized_hotkey(normalized_hotkey)
        pyautogui_keys = []
        for mod in modifiers:
            if mod == "ctrl":
                pyautogui_keys.append("ctrl")
            elif mod == "alt":
                pyautogui_keys.append("alt")
            elif mod == "shift":
                pyautogui_keys.append("shift")
        pyautogui_keys.append("space" if key == "space" else key)

        try:
            main_window.set_focus()
            time.sleep(0.2)
            main_window.type_keys(pywinauto_sequence, with_spaces=False, pause=0.1)
            shortcut_sent = True
            self.log("OK: Sent shortcut using pywinauto type_keys()")
        except Exception as e:
            last_error = str(e)
            self.log(f"  pywinauto type_keys() failed: {str(e)[:60]}")

        if not shortcut_sent:
            try:
                main_window.set_focus()
                time.sleep(0.2)
                main_window.send_keystrokes(pywinauto_sequence)
                shortcut_sent = True
                self.log("OK: Sent shortcut using pywinauto send_keystrokes()")
            except Exception as e:
                last_error = str(e)
                self.log(f"  pywinauto send_keystrokes() failed: {str(e)[:60]}")

        if not shortcut_sent:
            try:
                main_window.set_focus()
                time.sleep(0.2)
                pyautogui.hotkey(*pyautogui_keys)
                shortcut_sent = True
                self.log("OK: Sent shortcut using pyautogui hotkey()")
            except Exception as e:
                last_error = str(e)
                self.log(f"  pyautogui hotkey() failed: {str(e)[:60]}")

        if not shortcut_sent:
            try:
                main_window.set_focus()
                time.sleep(0.2)
                for mod in pyautogui_keys[:-1]:
                    pyautogui.keyDown(mod)
                pyautogui.press(pyautogui_keys[-1])
                for mod in reversed(pyautogui_keys[:-1]):
                    pyautogui.keyUp(mod)
                shortcut_sent = True
                self.log("OK: Sent shortcut using pyautogui keyDown/Up()")
            except Exception as e:
                last_error = str(e)
                self.log(f"  pyautogui keyDown/Up() failed: {str(e)[:60]}")

        return shortcut_sent, last_error

    def trigger_save_selection(self):
        if IS_WINDOWS:
            if not PYWINAUTO_AVAILABLE or not PYAUTOGUI_AVAILABLE:
                messagebox.showerror(
                    "Missing Dependencies",
                    "This feature requires pywinauto and pyautogui.\n\n"
                    "Install them with:\n"
                    "pip install pywinauto pyautogui",
                )
                return
        elif not IS_MACOS:
            messagebox.showerror("Platform Error", "This feature is only available on macOS and Windows.")
            return

        thread = threading.Thread(target=self._trigger_save_selection_thread, daemon=True)
        thread.start()

    def _trigger_save_selection_thread(self):
        if IS_MACOS:
            self._trigger_save_selection_macos()
        else:
            self._trigger_save_selection_windows()

    def _trigger_save_selection_macos(self):
        try:
            program_id = self._get_selected_program_id()
            program_label = self._get_program_label(program_id)
            effective_hotkey, hotkey_error = self._resolve_effective_hotkey(program_id)
            if not effective_hotkey:
                self.log(f"Error: {hotkey_error}")
                self.root.after(
                    0,
                    lambda: messagebox.showerror("Shortcut Not Configured", hotkey_error),
                )
                return

            self.log(f"Attempting to trigger save/export in {program_label}...")
            self.log(f"Step 1: Finding {program_label} window...")
            success, output, error = self._find_program_window_macos(program_id)
            if not success:
                self.log(f"Error: Could not find {program_label} window")
                self.log(f"Error: {error}")
                self.root.after(
                    0,
                    lambda: messagebox.showerror(
                        f"{program_label} Not Found",
                        f"Could not find a running {program_label} window.\n\n"
                        f"Please ensure {program_label} is open and try again.\n\n"
                        "Check the output log for details.",
                    ),
                )
                return

            self.log(f"OK: Found {program_label}: {output}")
            self.log(f"Step 2: Activating {program_label} window...")
            activated, _, activate_error = self._activate_program_window_macos(program_id)
            if activated:
                self.log(f"OK: Activated {program_label} window")
            else:
                self.log(f"Warning: Could not activate {program_label}: {activate_error}")

            time.sleep(0.5)
            shortcut_label = _format_hotkey_label(effective_hotkey)
            self.log(f"Step 3: Sending keyboard shortcut {shortcut_label}...")
            sent, _, send_error = self._send_hotkey_macos(effective_hotkey)
            if sent:
                self.log("OK: Keyboard shortcut sent successfully!")
                time.sleep(0.5)
                self.log(f"Please complete the save/export dialog in {program_label}...")
            else:
                self.log(f"Error: Failed to send keyboard shortcut: {send_error}")
                self.root.after(
                    0,
                    lambda: messagebox.showerror(
                        "Error",
                        f"Could not send keyboard shortcut {shortcut_label} to {program_label}.\n\n"
                        "Please check the output log for details.",
                    ),
                )
        except Exception as e:
            error_msg = f"Error triggering save/export: {str(e)}"
            self.log(f"Error: {error_msg}")
            import traceback

            self.log(f"Traceback:\n{traceback.format_exc()}")
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))

    def _trigger_save_selection_windows(self):
        app = None
        main_window = None

        try:
            program_id = self._get_selected_program_id()
            program_label = self._get_program_label(program_id)
            effective_hotkey, hotkey_error = self._resolve_effective_hotkey(program_id)
            if not effective_hotkey:
                self.log(f"Error: {hotkey_error}")
                self.root.after(
                    0,
                    lambda: messagebox.showerror("Shortcut Not Configured", hotkey_error),
                )
                return

            self.log(f"Attempting to trigger save/export in {program_label}...")
            self.log(f"Step 1: Finding {program_label} window...")
            app, methods_tried = self._find_program_window_windows(program_id)
            if app is None:
                self.log(f"Error: Could not find {program_label} window")
                self.log("Methods tried:")
                for method in methods_tried:
                    self.log(f"  - {method}")
                self.root.after(
                    0,
                    lambda: messagebox.showerror(
                        f"{program_label} Not Found",
                        f"Could not find a running {program_label} window.\n\n"
                        f"Please ensure {program_label} is open and try again.\n\n"
                        "Check the output log for details.",
                    ),
                )
                return

            self.log("Step 2: Accessing main window...")
            try:
                main_window = app.top_window()
                window_title = main_window.window_text()
                self.log(f"OK: Found window: '{window_title}'")
            except Exception as e:
                self.log(f"Error: Could not access {program_label} window: {str(e)}")
                self.root.after(
                    0,
                    lambda: messagebox.showerror("Error", f"Could not access {program_label} window: {str(e)}"),
                )
                return

            self.log(f"Step 3: Activating {program_label} window...")
            activated = False
            try:
                main_window.set_focus()
                time.sleep(0.5)
                activated = True
                self.log("OK: Activated using set_focus()")
            except Exception as e:
                self.log(f"  set_focus() failed: {str(e)[:60]}")

            if not activated:
                try:
                    main_window.set_foreground()
                    time.sleep(0.5)
                    activated = True
                    self.log("OK: Activated using set_foreground()")
                except Exception as e:
                    self.log(f"  set_foreground() failed: {str(e)[:60]}")

            if not activated and WIN32_AVAILABLE:
                try:
                    hwnd = main_window.handle
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
                    win32gui.BringWindowToTop(hwnd)
                    time.sleep(0.5)
                    activated = True
                    self.log("OK: Activated using Windows API")
                except Exception as e:
                    self.log(f"  Windows API activation failed: {str(e)[:60]}")

            if not activated:
                self.log("Warning: Could not activate window, continuing anyway...")

            time.sleep(0.5)
            shortcut_label = _format_hotkey_label(effective_hotkey)
            self.log(f"Step 4: Sending keyboard shortcut {shortcut_label}...")
            sent, send_error = self._send_hotkey_windows(main_window, effective_hotkey)
            if sent:
                self.log("OK: Keyboard shortcut sent successfully!")
                time.sleep(0.5)
                self.log(f"Please complete the save/export dialog in {program_label}...")
            else:
                self.log(f"Error: Failed to send keyboard shortcut with all methods ({send_error})")
                self.root.after(
                    0,
                    lambda: messagebox.showerror(
                        "Error",
                        f"Could not send keyboard shortcut {shortcut_label} to {program_label}.\n\n"
                        "Please check the output log for details.",
                    ),
                )

        except Exception as e:
            error_msg = f"Error triggering save/export: {str(e)}"
            self.log(f"Error: {error_msg}")
            import traceback

            self.log(f"Traceback:\n{traceback.format_exc()}")
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))

    def setup_hotkey_request_monitor(self):
        self.hotkey_request_path = HOTKEY_REQUEST_FILE
        try:
            if not self.hotkey_request_path.exists():
                self.hotkey_request_path.write_text("0")
            self._last_hotkey_request = self.hotkey_request_path.stat().st_mtime
        except Exception:
            self._last_hotkey_request = 0

        monitor_thread = threading.Thread(target=self._monitor_hotkey_request, daemon=True)
        monitor_thread.start()

    def _monitor_hotkey_request(self):
        while not self._hotkey_monitor_stop.is_set():
            try:
                if self.hotkey_request_path.exists():
                    mtime = self.hotkey_request_path.stat().st_mtime
                    if mtime > self._last_hotkey_request:
                        self._last_hotkey_request = mtime
                        self.log("Global hotkey request detected (background listener, triggering save/export).")
                        self.root.after(0, self.trigger_save_selection)
                time.sleep(0.6)
            except Exception:
                time.sleep(1)

    def register_global_hotkey(self):
        if self.disable_global_hotkey:
            self.log(
                f"Global hotkey registration skipped (external listener handles {_format_hotkey_label(GLOBAL_HOTKEY)})."
            )
            return

        hotkey = GLOBAL_HOTKEY

        if IS_MACOS:
            if not PYNPUT_AVAILABLE:
                self.log("Global hotkey not available: Install 'pynput' library (pip install pynput)")
                return
            pynput_hotkey = (
                hotkey.replace("cmd", "<cmd>")
                .replace("ctrl", "<ctrl>")
                .replace("alt", "<alt>")
                .replace("shift", "<shift>")
            )

            def on_activate():
                self.root.after(0, self.trigger_save_selection)

            def run_pynput_listener():
                with pynput_keyboard.GlobalHotKeys({pynput_hotkey: on_activate}) as listener:
                    self._pynput_listener = listener
                    listener.join()

            try:
                thread = threading.Thread(target=run_pynput_listener, daemon=True)
                thread.start()
                self.log(f"OK: Global hotkey registered: {_format_hotkey_label(hotkey)}")
                self.log("  You can now press the hotkey from anywhere to trigger save/export automation!")
            except Exception as e:
                self.log(f"Error: Failed to register global hotkey: {str(e)}")
                messagebox.showwarning(
                    "Hotkey Registration Failed",
                    f"Could not register global hotkey:\n{str(e)}\n\n"
                    "You can still use the button in the app.",
                )
            return

        if not KEYBOARD_AVAILABLE:
            self.log("Global hotkey not available: Install 'keyboard' library (pip install keyboard)")
            return

        try:
            keyboard.add_hotkey(hotkey, self.trigger_save_selection, suppress=False)
            self.log(f"OK: Global hotkey registered: {_format_hotkey_label(hotkey)}")
            self.log("  You can now press the hotkey from anywhere to trigger save/export automation!")
        except Exception as e:
            self.log(f"Error: Failed to register global hotkey: {str(e)}")
            messagebox.showwarning(
                "Hotkey Registration Failed",
                f"Could not register global hotkey:\n{str(e)}\n\n"
                "You can still use the button in the app.",
            )

    def on_closing(self):
        if IS_MACOS and self._pynput_listener is not None:
            try:
                self._pynput_listener.stop()
            except Exception:
                pass
        elif not IS_MACOS and KEYBOARD_AVAILABLE:
            try:
                keyboard.unhook_all_hotkeys()
            except Exception:
                pass
        self._hotkey_monitor_stop.set()
        watching_state = self.watching
        self.save_preferences(watching_override=watching_state)
        self.watching = False
        self.root.destroy()


def main():
    parser = argparse.ArgumentParser(description="MuseScore Pitch Extractor GUI")
    parser.add_argument(
        "--trigger-save-selection",
        action="store_true",
        help="Trigger the Save Selection automation as soon as the UI loads.",
    )
    parser.add_argument(
        "--disable-global-hotkey",
        action="store_true",
        help="Skip registering the in-app hotkey (used by the background listener).",
    )

    args = parser.parse_args()

    root = tk.Tk()
    app = MuseScoreExtractorApp(
        root,
        trigger_on_start=args.trigger_save_selection,
        disable_global_hotkey=args.disable_global_hotkey,
    )
    root.mainloop()


if __name__ == "__main__":
    main()
