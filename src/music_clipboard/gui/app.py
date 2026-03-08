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
import urllib.error
import urllib.request
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk

if __package__ is None or __package__ == "":
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from music_clipboard.platform.runtime import IS_MACOS, IS_WINDOWS, default_hotkey, output_dirs

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
TAB_CLIPBOARD = "clipboard"
TAB_AI_EDITING = "ai_editing"
TAB_SETTINGS = "settings"
AI_FLOW_CLAUDE = "claude"
AI_FLOW_OPENAI_MINIMAL = "openai_minimal"
AI_FLOW_LABELS = {
    AI_FLOW_CLAUDE: "Claude (MuseScore automation)",
    AI_FLOW_OPENAI_MINIMAL: "OpenAI MIDI (minimal)",
}
OPENAI_MODEL_CLASSIFIER = "gpt-5-nano"
OPENAI_MODEL_SEMANTIC = "gpt-5.1"
OPENAI_MODEL_MIDI = "gpt-5.2"
OPENAI_MIDI_CHUNK_THRESHOLD_CHARS = 180000
OPENAI_SEMANTIC_SENTINEL = "NO_SEMANTIC_QUESTION"
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
    from music_clipboard.extract.pitches_with_position import extract_pitches_with_position_from_mscx

    EXTRACTION_FUNCTION = extract_pitches_with_position_from_mscx
    EXTRACTION_SCRIPT = "extract_pitches_with_position"
except ImportError:
    try:
        from music_clipboard.extract.pitches import extract_pitches_from_mscx

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
    from music_clipboard.extract.midi import extract_midi_from_mscx

    MIDI_EXTRACTION_FUNCTION = extract_midi_from_mscx
except ImportError:
    MIDI_EXTRACTION_FUNCTION = None

try:
    import mido
    MIDO_AVAILABLE = True
except ImportError:
    mido = None
    MIDO_AVAILABLE = False


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
                "Please ensure extraction modules are importable:\n"
                "- music_clipboard.extract.pitches_with_position\n"
                "- music_clipboard.extract.pitches",
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
        self.output_views = []
        self.open_location_buttons = []
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
        self.ai_flow_var = tk.StringVar(value=AI_FLOW_LABELS[AI_FLOW_OPENAI_MINIMAL])

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
            "active_tab": TAB_CLIPBOARD,
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
        if merged.get("active_tab") not in {TAB_CLIPBOARD, TAB_AI_EDITING, TAB_SETTINGS}:
            merged["active_tab"] = TAB_CLIPBOARD

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
            "active_tab": self._get_active_tab_id(),
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
        self._select_tab_by_id(self.preferences.get("active_tab", TAB_CLIPBOARD))
        self._update_program_dependent_ui()

        if self.preferences.get("watching"):
            folder = self.watch_folder.get().strip()
            if folder and os.path.exists(folder):
                self.root.after(200, self.toggle_watch)

    def _get_active_tab_id(self):
        if not hasattr(self, "notebook"):
            return TAB_CLIPBOARD
        selected = self.notebook.select()
        if selected == str(self.ai_tab):
            return TAB_AI_EDITING
        if selected == str(self.settings_tab):
            return TAB_SETTINGS
        return TAB_CLIPBOARD

    def _select_tab_by_id(self, tab_id):
        if not hasattr(self, "notebook"):
            return
        if tab_id == TAB_AI_EDITING:
            self.notebook.select(self.ai_tab)
        elif tab_id == TAB_SETTINGS:
            self.notebook.select(self.settings_tab)
        else:
            self.notebook.select(self.clipboard_tab)

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

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.clipboard_tab = ttk.Frame(self.notebook, padding="10")
        self.ai_tab = ttk.Frame(self.notebook, padding="10")
        self.settings_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.clipboard_tab, text="Clipboard")
        self.notebook.add(self.ai_tab, text="AI Editing")
        self.notebook.add(self.settings_tab, text="Settings")

        self._build_clipboard_tab()
        self._build_ai_tab()
        self._build_settings_tab()

        self.selected_program_dropdown.bind("<<ComboboxSelected>>", self._on_selected_program_changed)
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self._refresh_program_dropdown()
        self._update_program_dependent_ui()

    def _build_output_panel(self, parent):
        output_frame = ttk.LabelFrame(parent, text="Output", padding="10")
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)

        output_text = scrolledtext.ScrolledText(output_frame, height=15, width=80, wrap=tk.WORD)
        output_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.output_views.append(output_text)

        button_frame = ttk.Frame(output_frame)
        button_frame.grid(row=1, column=0, pady=5)

        ttk.Button(button_frame, text="Clear Output", command=self.clear_output).grid(row=0, column=0, padx=5)

        open_location_button = ttk.Button(
            button_frame,
            text="Open File Location",
            command=self.open_file_location,
            state="disabled",
        )
        open_location_button.grid(row=0, column=1, padx=5)
        self.open_location_buttons.append(open_location_button)

        ttk.Checkbutton(
            button_frame,
            text="Auto-delete previous extraction",
            variable=self.delete_previous_var,
        ).grid(row=0, column=2, padx=5)

        return output_frame, output_text, open_location_button

    def _build_clipboard_tab(self):
        clipboard_tab = self.clipboard_tab
        clipboard_tab.columnconfigure(0, weight=1)
        clipboard_tab.rowconfigure(2, weight=1)

        title_label = ttk.Label(
            clipboard_tab,
            text="Music Clipboard Extractor",
            font=("Arial", 16, "bold"),
        )
        title_label.grid(row=0, column=0, pady=(0, 20), sticky=tk.W)

        watch_frame = ttk.LabelFrame(clipboard_tab, text="Auto-Process Saved Selections", padding="10")
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

        self.clipboard_instructions_label = ttk.Label(watch_frame, justify=tk.LEFT, foreground="gray")
        self.clipboard_instructions_label.grid(row=6, column=0, columnspan=3, pady=10, sticky=tk.W)

        output_frame, output_text, open_location_button = self._build_output_panel(clipboard_tab)
        output_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        self.output_text = output_text
        self.open_location_button = open_location_button

    def _build_ai_tab(self):
        ai_tab = self.ai_tab
        ai_tab.columnconfigure(0, weight=1)
        ai_tab.rowconfigure(2, weight=1)

        title_label = ttk.Label(
            ai_tab,
            text="AI Editing",
            font=("Arial", 16, "bold"),
        )
        title_label.grid(row=0, column=0, pady=(0, 20), sticky=tk.W)

        ai_frame = ttk.LabelFrame(ai_tab, text="AI Flow", padding="10")
        ai_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        ai_frame.columnconfigure(0, weight=1)

        self.ai_mode_label = ttk.Label(ai_frame, justify=tk.LEFT, foreground="gray")
        self.ai_mode_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 6))

        self.ai_instructions_label = ttk.Label(ai_frame, justify=tk.LEFT, foreground="gray")
        self.ai_instructions_label.grid(row=1, column=0, sticky=tk.W)

        ai_flow_frame = ttk.Frame(ai_frame)
        ai_flow_frame.grid(row=2, column=0, sticky=tk.W, pady=(10, 0))
        ttk.Label(ai_flow_frame, text="AI Option:").grid(row=0, column=0, sticky=tk.W, padx=(0, 6))
        self.ai_flow_dropdown = ttk.Combobox(
            ai_flow_frame,
            textvariable=self.ai_flow_var,
            state="readonly",
            values=list(AI_FLOW_LABELS.values()),
            width=34,
        )
        self.ai_flow_dropdown.grid(row=0, column=1, sticky=tk.W)
        self.ai_flow_var.set(AI_FLOW_LABELS[AI_FLOW_OPENAI_MINIMAL])
        self.ai_flow_dropdown.bind("<<ComboboxSelected>>", self._on_ai_flow_changed)

        output_frame, _, _ = self._build_output_panel(ai_tab)
        output_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

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

    def _on_tab_changed(self, _event=None):
        self._update_program_dependent_ui()
        self.save_preferences()

    def _on_ai_flow_changed(self, _event=None):
        self._update_program_dependent_ui()

    def _is_ai_editing_active(self):
        return self._get_active_tab_id() == TAB_AI_EDITING

    def _get_selected_ai_flow(self):
        selected_label = (self.ai_flow_var.get() or "").strip()
        for flow_id, label in AI_FLOW_LABELS.items():
            if selected_label == label:
                return flow_id
        return AI_FLOW_OPENAI_MINIMAL

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
            "\nMode: Clipboard (extraction/clipboard workflow only)"
        )

    def _build_ai_instruction_text(self):
        selected_program = self._get_selected_program_id()
        selected_label = self._get_program_label(selected_program)
        selected_flow = self._get_selected_ai_flow()
        selected_flow_label = AI_FLOW_LABELS.get(selected_flow, AI_FLOW_LABELS[AI_FLOW_CLAUDE])
        flow_notes = (
            "- Claude flow opens MuseScore + plugin and asks Claude to export MIDI."
            if selected_flow == AI_FLOW_CLAUDE
            else "- OpenAI flow sends MIDI as text JSON + your prompt, then rebuilds returned MIDI."
        )
        return (
            "AI Editing Mode:\n"
            "- New watched files still run extraction first.\n"
            "- AI automation runs only while this tab is active.\n"
            f"- Save/export trigger target program: {selected_label}\n"
            f"- Selected AI option: {selected_flow_label}\n"
            f"{flow_notes}"
        )

    def _update_program_dependent_ui(self):
        selected_program = self._get_selected_program_id()
        selected_label = self._get_program_label(selected_program)
        self.save_selection_button.config(text=f"Trigger Save/Export in {selected_label}")
        self.clipboard_instructions_label.config(text=self._build_instruction_text())
        ai_active = self._is_ai_editing_active()
        self.ai_mode_label.config(
            text=(
                "Status: AI Editing ACTIVE (automation enabled)"
                if ai_active
                else "Status: AI Editing inactive (select this tab to enable automation)"
            ),
            foreground=("green" if ai_active else "gray"),
        )
        self.ai_instructions_label.config(text=self._build_ai_instruction_text())

    def _parse_dotenv_key(self, env_name):
        project_root = Path(__file__).resolve().parents[3]
        env_path = project_root / ".env"
        if not env_path.exists():
            return ""
        try:
            for line in env_path.read_text(encoding="utf-8").splitlines():
                stripped = line.strip()
                if not stripped or stripped.startswith("#") or "=" not in stripped:
                    continue
                key, value = stripped.split("=", 1)
                normalized_key = key.strip()
                if normalized_key.startswith("export "):
                    normalized_key = normalized_key[len("export ") :].strip()
                if normalized_key != env_name:
                    continue
                cleaned = value.strip().strip('"').strip("'")
                return cleaned
        except Exception:
            return ""
        return ""

    def _resolve_openai_api_key(self):
        api_key = (os.environ.get("OPENAI_KEY") or "").strip()
        if api_key:
            return api_key
        return self._parse_dotenv_key("OPENAI_KEY")

    def _openai_request_json(self, url, api_key, payload):
        request = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        with urllib.request.urlopen(request, timeout=180) as response:
            return json.loads(response.read().decode("utf-8"))

    def _openai_responses_call(self, api_key, model, content_items):
        payload = {
            "model": model,
            "input": [{"role": "user", "content": content_items}],
        }
        return self._openai_request_json("https://api.openai.com/v1/responses", api_key, payload)

    def _extract_text_from_openai_response(self, response_json):
        output_text = (response_json or {}).get("output_text")
        if isinstance(output_text, str) and output_text.strip():
            return output_text.strip()

        output_items = (response_json or {}).get("output", [])
        fragments = []
        for item in output_items:
            content_items = item.get("content", []) if isinstance(item, dict) else []
            for content in content_items:
                if not isinstance(content, dict):
                    continue
                if content.get("type") in ("output_text", "text"):
                    text_value = content.get("text")
                    if isinstance(text_value, str) and text_value.strip():
                        fragments.append(text_value.strip())
        return "\n".join(fragments).strip()

    def _extract_json_object_from_text(self, response_text):
        raw_text = (response_text or "").strip()
        if not raw_text:
            return None

        if raw_text.startswith("```"):
            lines = raw_text.splitlines()
            if len(lines) >= 3 and lines[0].startswith("```") and lines[-1].startswith("```"):
                raw_text = "\n".join(lines[1:-1]).strip()

        payload = None
        try:
            payload = json.loads(raw_text)
        except json.JSONDecodeError:
            if "{" in raw_text and "}" in raw_text:
                start = raw_text.find("{")
                end = raw_text.rfind("}")
                if start != -1 and end > start:
                    try:
                        payload = json.loads(raw_text[start : end + 1])
                    except json.JSONDecodeError:
                        payload = None
        if not isinstance(payload, dict):
            return None

        return payload

    def _classify_prompt_for_semantic_and_edit_local(self, prompt_text):
        prompt_lower = (prompt_text or "").strip().lower()
        semantic_indicators = ["why", "how", "what is", "explain", "theory", "meaning"]
        edit_indicators = ["transpose", "quantize", "humanize", "velocity", "swing", "note", "bar", "chord"]
        has_semantic = any(token in prompt_lower for token in semantic_indicators)
        has_edit = any(token in prompt_lower for token in edit_indicators)
        return has_semantic, has_edit

    def _run_openai_classifier_call(self, api_key, prompt_text):
        content = [
            {
                "type": "input_text",
                "text": (
                    "Classify whether the user prompt includes a semantic/music-theory explanation request.\n"
                    "Return ONLY valid JSON with exactly:\n"
                    "{ \"needs_semantic\": <true|false> }\n"
                    "Set true if any semantic explanation is requested. Otherwise false."
                ),
            },
            {"type": "input_text", "text": f"User prompt:\n{prompt_text.strip()}"},
        ]
        response_json = self._openai_responses_call(api_key, OPENAI_MODEL_CLASSIFIER, content)
        response_text = self._extract_text_from_openai_response(response_json)
        response_obj = self._extract_json_object_from_text(response_text)
        if not isinstance(response_obj, dict) or set(response_obj.keys()) != {"needs_semantic"}:
            raise RuntimeError("Classifier did not return required JSON shape.")
        needs_semantic = response_obj.get("needs_semantic")
        if not isinstance(needs_semantic, bool):
            raise RuntimeError("Classifier 'needs_semantic' must be boolean.")
        return needs_semantic

    def _run_openai_semantic_call(self, api_key, prompt_text):
        content = [
            {
                "type": "input_text",
                "text": (
                    "You are for semantic music Q&A only.\n"
                    f"If there is no semantic question, return exactly {OPENAI_SEMANTIC_SENTINEL}.\n"
                    "If there is a semantic question, return ONLY valid JSON:\n"
                    "{ \"assistant_text\": \"<answer>\" }\n"
                    "Do not provide MIDI edits or event-level instructions."
                ),
            },
            {"type": "input_text", "text": f"User prompt:\n{prompt_text.strip()}"},
        ]
        response_json = self._openai_responses_call(api_key, OPENAI_MODEL_SEMANTIC, content)
        return self._extract_text_from_openai_response(response_json)

    def _parse_semantic_response_text(self, response_text):
        text = (response_text or "").strip()
        if not text:
            return None
        if text == OPENAI_SEMANTIC_SENTINEL:
            return None
        payload = self._extract_json_object_from_text(text)
        if isinstance(payload, dict):
            assistant_text = payload.get("assistant_text")
            if isinstance(assistant_text, str) and assistant_text.strip():
                return assistant_text.strip()
        return text

    def _ensure_midi_input_for_openai(self, file_path):
        extension = Path(file_path).suffix.lower()
        if extension in (".mid", ".midi"):
            return file_path, False

        if extension in EXTRACTABLE_SCORE_EXTENSIONS:
            if MIDI_EXTRACTION_FUNCTION is None:
                raise RuntimeError("MIDI extraction function is not available.")
            with tempfile.NamedTemporaryFile(suffix=".mid", delete=False) as tmp:
                temp_midi_path = tmp.name
            midi_path = MIDI_EXTRACTION_FUNCTION(file_path, temp_midi_path)
            if not midi_path or not os.path.exists(midi_path):
                raise RuntimeError("Could not prepare MIDI input from score.")
            return midi_path, True

        raise RuntimeError(f"Unsupported file type for OpenAI MIDI flow: {extension}")

    def _midi_to_text_payload(self, midi_path):
        if not MIDO_AVAILABLE:
            raise RuntimeError("mido is required for OpenAI MIDI text flow. Install with: pip install mido")

        midi_obj = mido.MidiFile(midi_path)
        tracks_payload = []
        for track in midi_obj.tracks:
            track_messages = []
            for msg in track:
                msg_dict = msg.dict()
                msg_dict["is_meta"] = bool(msg.is_meta)
                track_messages.append(msg_dict)
            tracks_payload.append({"name": track.name or "", "messages": track_messages})

        return {
            "type": int(midi_obj.type),
            "ticks_per_beat": int(midi_obj.ticks_per_beat),
            "tracks": tracks_payload,
        }

    def _validate_midi_payload(self, midi_payload):
        if not MIDO_AVAILABLE:
            raise RuntimeError("mido is required for OpenAI MIDI text flow. Install with: pip install mido")
        if not isinstance(midi_payload, dict):
            raise RuntimeError("Returned MIDI payload is not a JSON object.")

        ticks_per_beat = int(midi_payload.get("ticks_per_beat", 480))
        midi_type = int(midi_payload.get("type", 1))
        tracks = midi_payload.get("tracks")
        if not isinstance(tracks, list) or not tracks:
            raise RuntimeError("Returned MIDI JSON must include a non-empty 'tracks' list.")

        validated_tracks = []
        for track_obj in tracks:
            if not isinstance(track_obj, dict):
                continue
            out_track = mido.MidiTrack()
            out_track.name = (track_obj.get("name") or "").strip()
            for message_obj in track_obj.get("messages", []):
                if not isinstance(message_obj, dict):
                    continue
                msg_data = dict(message_obj)
                is_meta = bool(msg_data.pop("is_meta", False))
                msg_time = msg_data.get("time", 0)
                if int(msg_time) < 0:
                    raise RuntimeError("Invalid MIDI message returned by model: negative delta-time is not allowed.")
                try:
                    if is_meta:
                        msg = mido.MetaMessage.from_dict(msg_data)
                    else:
                        msg = mido.Message.from_dict(msg_data)
                except Exception as exc:
                    raise RuntimeError(f"Invalid MIDI message returned by model: {exc}")
                out_track.append(msg)
            validated_tracks.append(out_track)

        if not validated_tracks:
            raise RuntimeError("Returned MIDI JSON produced no valid tracks.")
        return midi_type, ticks_per_beat, validated_tracks

    def _text_payload_to_midi_file(self, midi_payload, output_path):
        midi_type, ticks_per_beat, validated_tracks = self._validate_midi_payload(midi_payload)
        out_midi = mido.MidiFile(type=midi_type, ticks_per_beat=ticks_per_beat)
        for track in validated_tracks:
            out_midi.tracks.append(track)
        out_midi.save(output_path)

    def _build_openai_midi_edit_content(self, prompt_text, midi_payload, chunk_index=None, chunk_count=None, retry_error=None):
        chunk_text = "single payload"
        if chunk_index is not None and chunk_count is not None:
            chunk_text = f"chunk {chunk_index + 1} of {chunk_count}"

        retry_text = ""
        if retry_error:
            retry_text = (
                "\nThe previous response was invalid. Fix it and return strictly valid JSON. "
                f"Previous validation error: {retry_error}"
            )

        return [
            {
                "type": "input_text",
                "text": (
                    "You are editing a MIDI represented as JSON.\n"
                    "Return ONLY valid JSON with this exact top-level shape:\n"
                    "{ \"midi_json\": {\"type\": <int>, \"ticks_per_beat\": <int>, "
                    "\"tracks\": [{\"name\": <string>, \"messages\": [<message dicts>]}]}}\n"
                    "Message dicts must be compatible with mido Message.from_dict / MetaMessage.from_dict.\n"
                    "Use non-negative integer delta-time field 'time'.\n"
                    f"Current processing scope: {chunk_text}.{retry_text}\n"
                    "No markdown and no extra keys."
                ),
            },
            {"type": "input_text", "text": f"User prompt:\n{prompt_text.strip()}"},
            {"type": "input_text", "text": "Input MIDI JSON:\n" + json.dumps(midi_payload, separators=(",", ":"))},
        ]

    def _call_openai_midi_editor(self, api_key, prompt_text, midi_payload, chunk_index=None, chunk_count=None):
        last_error = ""
        for attempt in range(2):
            content = self._build_openai_midi_edit_content(
                prompt_text,
                midi_payload,
                chunk_index=chunk_index,
                chunk_count=chunk_count,
                retry_error=last_error if attempt == 1 else None,
            )
            response_json = self._openai_responses_call(api_key, OPENAI_MODEL_MIDI, content)
            response_text = self._extract_text_from_openai_response(response_json)
            if response_text:
                scope = (
                    f"chunk {chunk_index + 1}/{chunk_count}"
                    if chunk_index is not None and chunk_count is not None
                    else "single payload"
                )
                self.log(f"OpenAI MIDI response text ({scope}):")
                self.log(response_text)
            response_obj = self._extract_json_object_from_text(response_text)
            if not response_obj:
                last_error = "Response did not contain valid JSON."
                continue
            if set(response_obj.keys()) != {"midi_json"}:
                last_error = "Response JSON must contain only top-level key 'midi_json'."
                continue
            midi_output_payload = response_obj.get("midi_json")
            if not isinstance(midi_output_payload, dict):
                last_error = "Response JSON did not include top-level 'midi_json'."
                continue
            try:
                self._validate_midi_payload(midi_output_payload)
                return midi_output_payload
            except Exception as exc:
                last_error = str(exc)
                continue
        raise RuntimeError(f"OpenAI MIDI response invalid after retry: {last_error}")

    def _chunk_midi_payload_by_track(self, midi_payload, max_chars):
        tracks = midi_payload.get("tracks") or []
        if not tracks:
            return [midi_payload]

        chunks = []
        current = []
        for track in tracks:
            candidate = current + [track]
            candidate_payload = {
                "type": int(midi_payload.get("type", 1)),
                "ticks_per_beat": int(midi_payload.get("ticks_per_beat", 480)),
                "tracks": candidate,
            }
            candidate_len = len(json.dumps(candidate_payload, separators=(",", ":")))
            if current and candidate_len > max_chars:
                chunks.append(current)
                current = [track]
            else:
                current = candidate
        if current:
            chunks.append(current)

        chunk_payloads = []
        for chunk_tracks in chunks:
            chunk_payloads.append(
                {
                    "type": int(midi_payload.get("type", 1)),
                    "ticks_per_beat": int(midi_payload.get("ticks_per_beat", 480)),
                    "tracks": chunk_tracks,
                }
            )
        return chunk_payloads

    def _run_openai_midi_edit_flow(self, file_path, prompt_text):
        temporary_input = False
        midi_input_path = None
        try:
            api_key = self._resolve_openai_api_key()
            if not api_key:
                raise RuntimeError("OPENAI_KEY was not found in environment or .env.")

            midi_input_path, temporary_input = self._ensure_midi_input_for_openai(file_path)
            self.log(f"Preparing OpenAI MIDI edit for: {os.path.basename(midi_input_path)}")
            midi_text_payload = self._midi_to_text_payload(midi_input_path)
            has_edit = self._classify_prompt_for_semantic_and_edit_local(prompt_text)[1]
            try:
                needs_semantic = self._run_openai_classifier_call(api_key, prompt_text)
                self.log(
                    "Semantic gate: classifier "
                    f"({OPENAI_MODEL_CLASSIFIER}) -> {'run Call A' if needs_semantic else 'skip Call A'}."
                )
            except Exception as classifier_exc:
                needs_semantic = self._classify_prompt_for_semantic_and_edit_local(prompt_text)[0]
                self.log(
                    "Semantic gate: classifier failed; using local rules "
                    f"({'run Call A' if needs_semantic else 'skip Call A'}). Details: {classifier_exc}"
                )

            if needs_semantic:
                self.log(f"Running Call A (semantic) with {OPENAI_MODEL_SEMANTIC}.")
                try:
                    semantic_response = self._run_openai_semantic_call(api_key, prompt_text)
                    semantic_text = self._parse_semantic_response_text(semantic_response)
                    if semantic_text:
                        self.log("Call A assistant text:")
                        self.log(semantic_text)
                    else:
                        self.log("Call A returned no semantic answer.")
                except Exception as semantic_exc:
                    self.log(f"Call A failed; continuing with Call B. Details: {semantic_exc}")
            else:
                if has_edit:
                    self.log("Semantic gate: edit-only indicators found; skipping Call A.")
                else:
                    self.log("Semantic gate: no semantic indicators; skipping Call A.")

            serialized_len = len(json.dumps(midi_text_payload, separators=(",", ":")))
            if serialized_len > OPENAI_MIDI_CHUNK_THRESHOLD_CHARS:
                chunk_payloads = self._chunk_midi_payload_by_track(midi_text_payload, OPENAI_MIDI_CHUNK_THRESHOLD_CHARS)
                self.log(
                    f"Call B ({OPENAI_MODEL_MIDI}) processing mode: chunked by track "
                    f"({len(chunk_payloads)} chunks, payload chars={serialized_len})."
                )

                merged_tracks = []
                merged_type = int(midi_text_payload.get("type", 1))
                merged_tpb = int(midi_text_payload.get("ticks_per_beat", 480))
                total_chunks = len(chunk_payloads)
                for chunk_index, chunk_payload in enumerate(chunk_payloads):
                    self.log(f"Call B processing chunk {chunk_index + 1}/{total_chunks}...")
                    chunk_result = self._call_openai_midi_editor(
                        api_key,
                        prompt_text,
                        chunk_payload,
                        chunk_index=chunk_index,
                        chunk_count=total_chunks,
                    )
                    if chunk_index == 0:
                        merged_type = int(chunk_result.get("type", merged_type))
                        merged_tpb = int(chunk_result.get("ticks_per_beat", merged_tpb))
                    chunk_tracks = chunk_result.get("tracks") if isinstance(chunk_result, dict) else None
                    if not isinstance(chunk_tracks, list):
                        raise RuntimeError(f"Chunk {chunk_index + 1} did not return a valid 'tracks' list.")
                    merged_tracks.extend(chunk_tracks)

                midi_output_payload = {
                    "type": merged_type,
                    "ticks_per_beat": merged_tpb,
                    "tracks": merged_tracks,
                }
                self._validate_midi_payload(midi_output_payload)
            else:
                self.log(f"Call B ({OPENAI_MODEL_MIDI}) processing mode: single payload.")
                midi_output_payload = self._call_openai_midi_editor(api_key, prompt_text, midi_text_payload)

            os.makedirs(MIDI_OUTPUT_DIR, exist_ok=True)
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            output_path = os.path.join(MIDI_OUTPUT_DIR, f"{base_name}_openai_ai.mid")
            self._text_payload_to_midi_file(midi_output_payload, output_path)

            self.log(f"OpenAI MIDI edit saved to: {output_path}")
            self._handle_successful_extraction(output_path)
        except urllib.error.HTTPError as exc:
            details = ""
            try:
                details = exc.read().decode("utf-8")
            except Exception:
                details = str(exc)
            error_msg = f"OpenAI API request failed ({exc.code}): {details}"
            self.log(f"Error: {error_msg}")
            self._show_error_async("OpenAI Request Failed", error_msg)
        except Exception as exc:
            error_msg = str(exc)
            self.log(f"Error: {error_msg}")
            self._show_error_async("OpenAI MIDI Flow Failed", error_msg)
        finally:
            if temporary_input and midi_input_path and os.path.exists(midi_input_path):
                try:
                    os.remove(midi_input_path)
                except OSError:
                    pass

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
        targets = self.output_views if self.output_views else [self.output_text]
        for view in targets:
            view.insert(tk.END, message + "\n")
            view.see(tk.END)
        self.root.update_idletasks()

    def clear_output(self):
        targets = self.output_views if self.output_views else [self.output_text]
        for view in targets:
            view.delete(1.0, tk.END)

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
        self.root.after(
            0,
            lambda: [btn.config(state="normal") for btn in (self.open_location_buttons or [self.open_location_button])],
        )
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

        if not self._is_ai_editing_active():
            self.log(f"Clipboard mode: skipping AI automation for {os.path.basename(file_path)}")
            return

        prompt = simpledialog.askstring(
            "AI Edit Prompt",
            (
                f"Enter the AI edit prompt for:\n{os.path.basename(file_path)}\n\n"
                "The app will append your text to the selected AI option prompt."
            ),
            parent=self.root,
        )

        if prompt is None or not prompt.strip():
            self.log(f"Skipped AI flow for {os.path.basename(file_path)} (empty/canceled prompt).")
            return

        selected_flow = self._get_selected_ai_flow()
        if selected_flow == AI_FLOW_OPENAI_MINIMAL:
            thread = threading.Thread(
                target=self._run_openai_midi_edit_flow,
                args=(file_path, prompt.strip()),
                daemon=True,
            )
            thread.start()
            return

        if not IS_MACOS:
            self.log("Claude AI flow is only supported on macOS.")
            return
        thread = threading.Thread(target=self._run_ai_edit_flow, args=(file_path, prompt.strip()), daemon=True)
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
                        "MIDI extraction function not available. Please ensure music_clipboard.extract.midi is importable."
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
