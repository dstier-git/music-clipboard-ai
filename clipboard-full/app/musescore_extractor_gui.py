import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

if __package__ is None or __package__ == "":
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app.platform_utils import IS_MACOS, IS_WINDOWS, default_hotkey, output_dirs, save_selection_shortcut_label

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

OUTPUT_DIR, MIDI_OUTPUT_DIR = output_dirs()

GLOBAL_HOTKEY = default_hotkey()
SAVE_SELECTION_SHORTCUT_LABEL = save_selection_shortcut_label()

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
    parts = []
    for part in hotkey.split("+"):
        if part == "cmd":
            parts.append("Cmd")
        else:
            parts.append(part.capitalize())
    return "+".join(parts)


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
        self.output_format = tk.StringVar(value="Text")
        self.last_extracted_file = None
        self.delete_previous_var = tk.BooleanVar(value=True)
        self.preferences = self.load_preferences()
        self._hotkey_monitor_stop = threading.Event()
        self._last_hotkey_request = 0
        self._pynput_listener = None

        self.create_widgets()
        self.apply_saved_preferences()
        self.setup_hotkey_request_monitor()
        self.register_global_hotkey()

        if self.trigger_on_start:
            self.root.after(500, self.trigger_save_selection)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def load_preferences(self):
        if not CONFIG_FILE.exists():
            return {}
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                lines = [line.rstrip("\n") for line in f.readlines()]
            folder = lines[0] if lines else ""
            watching = False
            if len(lines) > 1:
                watching = lines[1].strip().lower() == "true"
            return {"watch_folder": folder, "watching": watching}
        except Exception:
            return {}

    def save_preferences(self, watching_override=None):
        folder = self.watch_folder.get()
        watching = self.watching if watching_override is None else watching_override
        self.preferences = {"watch_folder": folder, "watching": watching}
        try:
            CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                f.write(f"{folder}\n")
                f.write("true\n" if watching else "false\n")
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

        if self.preferences.get("watching"):
            folder = self.watch_folder.get().strip()
            if folder and os.path.exists(folder):
                self.root.after(200, self.toggle_watch)

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        title_label = ttk.Label(
            main_frame,
            text="MuseScore Pitch & Position Extractor",
            font=("Arial", 16, "bold"),
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="Select MuseScore File:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, width=50)
        file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(file_frame, text="Browse...", command=self.browse_file).grid(row=0, column=2, padx=5)
        ttk.Button(file_frame, text="Extract", command=self.extract_file).grid(row=0, column=3, padx=5)

        format_frame = ttk.Frame(file_frame)
        format_frame.grid(row=1, column=0, columnspan=4, sticky=tk.W, pady=(10, 0))
        ttk.Label(format_frame, text="Output Format:").grid(row=0, column=0, padx=5)
        format_dropdown = ttk.Combobox(
            format_frame,
            textvariable=self.output_format,
            values=["Text", "MIDI"],
            state="readonly",
        )
        format_dropdown.current(0)
        format_dropdown.grid(row=0, column=1, padx=5)

        # Replace with buttons instead:
            # ttk.Radiobutton(format_frame, text="Text", variable=self.output_format, value="text").grid(
            #     row=0, column=1, padx=5
            # )
            # ttk.Radiobutton(format_frame, text="MIDI", variable=self.output_format, value="midi").grid(
            #     row=0, column=2, padx=5
            # )

        watch_frame = ttk.LabelFrame(main_frame, text="Auto-Process Saved Selections", padding="10")
        watch_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        watch_frame.columnconfigure(1, weight=1)

        ttk.Label(watch_frame, text="Watch Folder:").grid(row=0, column=0, sticky=tk.W, padx=5)
        watch_entry = ttk.Entry(watch_frame, textvariable=self.watch_folder, width=50)
        watch_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(watch_frame, text="Browse...", command=self.browse_watch_folder).grid(row=0, column=2, padx=5)

        automation_frame = ttk.Frame(watch_frame)
        automation_frame.grid(row=1, column=0, columnspan=3, pady=5, sticky=tk.W)

        self.save_selection_button = ttk.Button(
            automation_frame,
            text="Trigger Save Selection in MuseScore",
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
        self.watch_button.grid(row=2, column=0, columnspan=3, pady=10)

        self.watch_status_label = ttk.Label(watch_frame, text="Status: Not watching", foreground="gray")
        self.watch_status_label.grid(row=3, column=0, columnspan=3)

        instructions = f"""
Instructions:
1. Manual Mode:
   - Select a .mscx or .mscz file
   - Click 'Extract' to process
2. Auto Mode (Save Selection):
   - Set the watch folder (where MuseScore saves selections), click 'Start Watching'
   - In MuseScore: Select the measures you want to extract
   - Click 'Trigger Save Selection in MuseScore' button (or manually: File > Save Selection)
   - Save in the watch folder
   - Shortcut reminder: {SAVE_SELECTION_SHORTCUT_LABEL}
        """
        ttk.Label(watch_frame, text=instructions.strip(), justify=tk.LEFT, foreground="gray").grid(
            row=4, column=0, columnspan=3, pady=10, sticky=tk.W
        )

        output_frame = ttk.LabelFrame(main_frame, text="Output", padding="10")
        output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)

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

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select MuseScore File",
            filetypes=[("MuseScore files", "*.mscx *.mscz"), ("All files", "*.*")],
        )
        if filename:
            self.file_path_var.set(filename)

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

    def _move_to_output_dir(self, src_path, output_ext):
        """
        Move a file into the app's output directory for the selected format.
        Returns the destination path on success, None on failure.
        Uses unique naming (e.g. "name (1).txt") if the destination already exists.
        """
        if not src_path or not os.path.exists(src_path):
            return None
        dest_dir = Path(MIDI_OUTPUT_DIR) if output_ext == ".mid" else Path(OUTPUT_DIR)
        try:
            dest_dir.mkdir(parents=True, exist_ok=True)
            name = os.path.basename(src_path)
            dest_path = dest_dir / name
            if dest_path.exists():
                stem, ext = os.path.splitext(name)
                n = 1
                while (dest_dir / f"{stem} ({n}){ext}").exists():
                    n += 1
                dest_path = dest_dir / f"{stem} ({n}){ext}"
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

    def extract_file(self, file_path=None):
        if file_path is None:
            file_path = self.file_path_var.get().strip().strip('"').strip("'")

        if not file_path:
            messagebox.showwarning("Warning", "Please select a file first.")
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
            if file.endswith((".mscx", ".mscz")):
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
                full_path = os.path.join(folder, file)
                initial_output_files.add(full_path)
        self.seen_output_type_files.update(initial_output_files)

        while self.watching:
            try:
                current_files = set()
                for file in os.listdir(folder):
                    if file.endswith((".mscx", ".mscz")):
                        full_path = os.path.join(folder, file)
                        current_files.add(full_path)

                        if full_path not in self.processed_files:
                            time.sleep(0.5)

                            try:
                                mod_time = os.path.getmtime(full_path)
                                if time.time() - mod_time > 1:
                                    self.processed_files.add(full_path)
                                    self.root.after(0, lambda f=full_path: self.extract_file(f))
                                    self.log(f"Detected new file: {os.path.basename(full_path)}")
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
                        full_path = os.path.join(folder, file)
                        current_output_files.add(full_path)
                        if full_path not in self.seen_output_type_files:
                            dest_path = self._move_to_output_dir(full_path, output_ext)
                            self.seen_output_type_files.add(full_path)
                            self.root.after(0, self._bring_app_to_front)
                            if dest_path:
                                self.root.after(0, lambda p=dest_path: self._reveal_file_in_folder(p))
                                self.log(f"Moved {os.path.basename(full_path)} to output folder: {dest_path}")
                            else:
                                self.log(f"Failed to move {os.path.basename(full_path)} to output folder")

                self.seen_output_type_files.intersection_update(current_output_files)
                time.sleep(1)

            except Exception as e:
                if self.watching:
                    self.log(f"Error watching folder: {str(e)}\n")
                time.sleep(2)

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
            self.log("Attempting to trigger Save Selection in MuseScore...")
            self.log("Step 1: Finding MuseScore window...")

            success, output, error = find_musescore_window_macos()
            if not success:
                self.log("Error: Could not find MuseScore window")
                self.log(f"Error: {error}")

                self.log("\nDebug: Searching for MuseScore processes...")
                found_processes = []

                if PSUTIL_AVAILABLE:
                    try:
                        for proc in psutil.process_iter(["pid", "name"]):
                            try:
                                proc_name = proc.info.get("name") or ""
                                proc_lower = proc_name.lower()
                                if proc_lower == "mscore" or "musescore" in proc_lower:
                                    found_processes.append(proc_name)
                                    self.log(
                                        f"  Found process: {proc_name} (PID: {proc.info.get('pid')})"
                                    )
                            except (psutil.NoSuchProcess, psutil.AccessDenied):
                                continue
                    except Exception as e:
                        self.log(f"  Could not list processes: {e}")

                list_script = """
                tell application "System Events"
                    set processList to name of every process
                    set resultList to {}
                    repeat with procName in processList
                        if procName contains "Muse" or procName contains "muse" or procName contains "Score" or procName contains "score" then
                            set end of resultList to procName
                        end if
                    end repeat
                    return resultList
                end tell
                """
                list_success, list_output, list_error = run_applescript(list_script)
                if list_success and list_output:
                    self.log(f"  AppleScript found processes: {list_output}")

                if not found_processes and not list_output:
                    self.log("  No MuseScore-related processes found.")
                    self.log("\nTroubleshooting tips:")
                    self.log("  1. Make sure MuseScore 4 is running with a score open")
                    self.log("  2. Check System Settings > Privacy & Security > Accessibility")
                    self.log("     - Ensure Terminal/Python has accessibility permissions")
                    self.log("  3. Try restarting MuseScore 4")

                self.root.after(
                    0,
                    lambda: messagebox.showerror(
                        "MuseScore Not Found",
                        "Could not find a running MuseScore window.\n\n"
                        "Please ensure MuseScore 4 is open with a score loaded and try again.\n\n"
                        "Check the output log for debugging information.\n\n"
                        "Note: You may need to grant accessibility permissions to Terminal/Python\n"
                        "in System Settings > Privacy & Security > Accessibility.",
                    ),
                )
                return

            self.log(f"OK: Found MuseScore: {output}")

            self.log("Step 2: Activating MuseScore window...")
            success, output, error = activate_musescore_window_macos()
            if success:
                self.log("OK: Activated MuseScore window")
            else:
                self.log(f"Warning: Warning: Could not activate window: {error}")

            time.sleep(0.5)

            self.log(f"Step 3: Sending keyboard shortcut {SAVE_SELECTION_SHORTCUT_LABEL}...")
            success, output, error = send_shortcut_macos()
            if success:
                self.log("OK: Keyboard shortcut sent successfully!")
                time.sleep(0.5)
                self.log("Please complete the save dialog in MuseScore...")
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Save Selection Triggered",
                        "Save Selection dialog should now be open in MuseScore.\n\n"
                        "Please:\n"
                        "1. Choose the save location (preferably the watch folder)\n"
                        "2. Enter a filename\n"
                        "3. Click Save\n\n"
                        "If watching is enabled, the file will be processed automatically.\n\n"
                        "If the dialog didn't open, check the output log for details.",
                    ),
                )
            else:
                self.log(f"Error: Failed to send keyboard shortcut: {error}")
                self.root.after(
                    0,
                    lambda: messagebox.showerror(
                        "Error",
                        "Could not send keyboard shortcut to MuseScore.\n\n"
                        "Please check the output log for details.\n\n"
                        "You can still use the manual method:\n"
                        f"File > Save Selection (or {SAVE_SELECTION_SHORTCUT_LABEL})",
                    ),
                )

        except Exception as e:
            error_msg = f"Error triggering Save Selection: {str(e)}"
            self.log(f"Error: {error_msg}")
            import traceback

            self.log(f"Traceback:\n{traceback.format_exc()}")
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))

    def _trigger_save_selection_windows(self):
        app = None
        main_window = None

        try:
            self.log("Attempting to trigger Save Selection in MuseScore...")
            self.log("Step 1: Finding MuseScore window...")

            found = False
            methods_tried = []

            try:
                app = Application(backend="uia").connect(path="MuseScore4.exe")
                methods_tried.append("UIA backend by process")
                found = True
                self.log("OK: Found MuseScore process (UIA backend by process name)")
            except Exception as e:
                methods_tried.append(f"UIA by process failed: {str(e)[:50]}")

            if not found:
                try:
                    app = Application(backend="win32").connect(path="MuseScore4.exe")
                    methods_tried.append("Win32 backend by process")
                    found = True
                    self.log("OK: Found MuseScore process (Win32 backend by process name)")
                except Exception as e:
                    methods_tried.append(f"Win32 by process failed: {str(e)[:50]}")

            if not found:
                try:
                    app = Application().connect(path="MuseScore4.exe")
                    methods_tried.append("Default backend by process")
                    found = True
                    self.log("OK: Found MuseScore process (default backend by process name)")
                except Exception as e:
                    methods_tried.append(f"Default by process failed: {str(e)[:50]}")

            if not found:
                try:
                    try:
                        app = Application(backend="uia").connect(title_re=".*MuseScore 4.*")
                        found = True
                        self.log("OK: Found MuseScore window (UIA backend by 'MuseScore 4' title)")
                    except Exception:
                        if PSUTIL_AVAILABLE:
                            for proc in psutil.process_iter(["pid", "name"]):
                                try:
                                    proc_name = proc.info["name"] or ""
                                    if "MuseScore4.exe" in proc_name or (
                                        "MuseScore" in proc_name and "Pitch" not in proc_name
                                    ):
                                        app = Application(backend="uia").connect(process=proc.info["pid"])
                                        found = True
                                        self.log(f"OK: Found MuseScore by process enumeration: {proc_name}")
                                        break
                                except (psutil.NoSuchProcess, psutil.AccessDenied):
                                    continue
                except Exception as e:
                    methods_tried.append(f"UIA by filtered title failed: {str(e)[:50]}")

            if not found and WIN32_AVAILABLE:
                try:
                    def enum_handler(hwnd, ctx):
                        if win32gui.IsWindowVisible(hwnd):
                            window_text = win32gui.GetWindowText(hwnd)
                            if "MuseScore" in window_text and "Pitch Extractor" not in window_text:
                                try:
                                    _, pid = win32gui.GetWindowThreadProcessId(hwnd)
                                    if PSUTIL_AVAILABLE:
                                        try:
                                            proc = psutil.Process(pid)
                                            proc_name = proc.name()
                                            if "MuseScore" in proc_name and "Pitch" not in proc_name:
                                                ctx.append((hwnd, window_text, proc_name))
                                        except Exception:
                                            if "MuseScore 4" in window_text or window_text.startswith("MuseScore"):
                                                ctx.append((hwnd, window_text, "unknown"))
                                    else:
                                        if "MuseScore 4" in window_text or window_text.startswith("MuseScore"):
                                            ctx.append((hwnd, window_text, "unknown"))
                                except Exception:
                                    pass

                    windows = []
                    win32gui.EnumWindows(enum_handler, windows)

                    if windows:
                        hwnd, window_text, proc_name = windows[0]
                        try:
                            app = Application().connect(handle=hwnd)
                            found = True
                            self.log(
                                f"OK: Found MuseScore window using Windows API: '{window_text}' (process: {proc_name})"
                            )
                        except Exception:
                            pass
                except Exception as e:
                    methods_tried.append(f"Windows API failed: {str(e)[:50]}")

            if not found and PSUTIL_AVAILABLE:
                try:
                    for proc in psutil.process_iter(["pid", "name", "exe"]):
                        try:
                            proc_name = proc.info["name"] or ""
                            proc_exe = proc.info["exe"] or ""
                            if "MuseScore4.exe" in proc_exe or (
                                proc_name and "MuseScore" in proc_name and "4" in proc_name and "Pitch" not in proc_name
                            ):
                                app = Application(backend="uia").connect(process=proc.info["pid"])
                                found = True
                                self.log(f"OK: Found MuseScore by process search: {proc_name}")
                                break
                        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                            continue
                except Exception as e:
                    methods_tried.append(f"Process search failed: {str(e)[:50]}")

            if not found:
                self.log("Error: Could not find MuseScore window")
                self.log("Methods tried:")
                for method in methods_tried:
                    self.log(f"  - {method}")
                self.root.after(
                    0,
                    lambda: messagebox.showerror(
                        "MuseScore Not Found",
                        "Could not find a running MuseScore window.\n\n"
                        "Please ensure MuseScore 4 is open with a score loaded and try again.\n\n"
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
                self.log(f"Error: Error accessing window: {str(e)}")
                self.root.after(
                    0,
                    lambda: messagebox.showerror("Error", f"Could not access MuseScore window: {str(e)}"),
                )
                return

            self.log("Step 3: Activating MuseScore window...")
            activated = False

            try:
                main_window.set_focus()
                time.sleep(0.5)
                activated = True
                self.log("OK: Activated using set_focus()")
            except Exception as e:
                self.log(f"  set_focus() failed: {str(e)[:50]}")

            if not activated:
                try:
                    main_window.set_foreground()
                    time.sleep(0.5)
                    activated = True
                    self.log("OK: Activated using set_foreground()")
                except Exception as e:
                    self.log(f"  set_foreground() failed: {str(e)[:50]}")

            if not activated:
                try:
                    main_window.restore()
                    main_window.set_focus()
                    time.sleep(0.5)
                    activated = True
                    self.log("OK: Activated using restore() + set_focus()")
                except Exception as e:
                    self.log(f"  restore() + set_focus() failed: {str(e)[:50]}")

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
                    self.log(f"  Windows API activation failed: {str(e)[:50]}")

            if not activated:
                self.log("Warning: Warning: Could not activate window, but continuing anyway...")

            time.sleep(0.5)

            self.log(f"Step 4: Sending keyboard shortcut {SAVE_SELECTION_SHORTCUT_LABEL}...")
            shortcut_sent = False

            try:
                main_window.set_focus()
                time.sleep(0.2)
                main_window.type_keys("^+s", with_spaces=False, pause=0.1)
                shortcut_sent = True
                self.log("OK: Sent shortcut using pywinauto type_keys()")
            except Exception as e:
                self.log(f"  pywinauto type_keys() failed: {str(e)[:50]}")

            if not shortcut_sent:
                try:
                    main_window.set_focus()
                    time.sleep(0.2)
                    main_window.send_keystrokes("^+s")
                    shortcut_sent = True
                    self.log("OK: Sent shortcut using pywinauto send_keystrokes()")
                except Exception as e:
                    self.log(f"  pywinauto send_keystrokes() failed: {str(e)[:50]}")

            if not shortcut_sent:
                try:
                    main_window.set_focus()
                    time.sleep(0.2)
                    pyautogui.hotkey("ctrl", "shift", "s")
                    shortcut_sent = True
                    self.log("OK: Sent shortcut using pyautogui hotkey()")
                except Exception as e:
                    self.log(f"  pyautogui hotkey() failed: {str(e)[:50]}")

            if not shortcut_sent:
                try:
                    main_window.set_focus()
                    time.sleep(0.2)
                    pyautogui.keyDown("ctrl")
                    pyautogui.keyDown("shift")
                    pyautogui.press("s")
                    pyautogui.keyUp("shift")
                    pyautogui.keyUp("ctrl")
                    shortcut_sent = True
                    self.log("OK: Sent shortcut using pyautogui keyDown/Up()")
                except Exception as e:
                    self.log(f"  pyautogui keyDown/Up() failed: {str(e)[:50]}")

            if shortcut_sent:
                self.log("OK: Keyboard shortcut sent successfully!")
                time.sleep(0.5)
                self.log("Please complete the save dialog in MuseScore...")
            else:
                self.log("Error: Failed to send keyboard shortcut with all methods")
                self.root.after(
                    0,
                    lambda: messagebox.showerror(
                        "Error",
                        "Could not send keyboard shortcut to MuseScore.\n\n"
                        "Please check the output log for details.\n\n"
                        "You can still use the manual method:\n"
                        f"File > Save Selection (or {SAVE_SELECTION_SHORTCUT_LABEL})",
                    ),
                )

        except Exception as e:
            error_msg = f"Error triggering Save Selection: {str(e)}"
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
                        self.log("Global hotkey request detected (background listener).")
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
                self.log("  You can now press the hotkey from anywhere to trigger Save Selection!")
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
            self.log("  You can now press the hotkey from anywhere to trigger Save Selection!")
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
