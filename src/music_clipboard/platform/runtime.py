import os
import platform
from pathlib import Path
from typing import List, Tuple

IS_WINDOWS = os.name == "nt"
IS_MACOS = platform.system() == "Darwin"


def project_root() -> Path:
    # /repo/src/music_clipboard/platform/runtime.py -> /repo
    return Path(__file__).resolve().parents[3]


def legacy_output_dirs() -> Tuple[Path, Path]:
    root = project_root()
    return root / "clipboard-full" / "txts", root / "clipboard-full" / "midis"


def output_dirs() -> Tuple[Path, Path]:
    root = project_root()
    text_dir = root / "data" / "outputs" / "text"
    midi_dir = root / "data" / "outputs" / "midi"
    return text_dir, midi_dir


def output_read_dirs() -> Tuple[List[Path], List[Path]]:
    text_dir, midi_dir = output_dirs()
    legacy_text_dir, legacy_midi_dir = legacy_output_dirs()

    text_candidates = [text_dir]
    midi_candidates = [midi_dir]

    if legacy_text_dir != text_dir and legacy_text_dir.exists():
        text_candidates.append(legacy_text_dir)
    if legacy_midi_dir != midi_dir and legacy_midi_dir.exists():
        midi_candidates.append(legacy_midi_dir)

    return text_candidates, midi_candidates


def default_hotkey() -> str:
    return "ctrl+cmd+s" if IS_MACOS else "ctrl+alt+s"


def save_selection_shortcut_label() -> str:
    return "Cmd+Shift+S" if IS_MACOS else "Ctrl+Shift+S"
