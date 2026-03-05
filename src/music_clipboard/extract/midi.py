import os
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

if __package__ is None or __package__ == "":
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from music_clipboard.platform.runtime import IS_MACOS, IS_WINDOWS, output_dirs

DEFAULT_MIDI_TEMPO = 120
MIDI_OUTPUT_DIR = output_dirs()[1]


def _find_musescore_exe():
    musescore_paths = []
    if IS_MACOS:
        musescore_paths = [
            "/Applications/MuseScore 4.app/Contents/MacOS/mscore",
            "/Applications/MuseScore 4.app/Contents/MacOS/MuseScore4",
            "/Applications/MuseScore.app/Contents/MacOS/mscore",
        ]
    elif IS_WINDOWS:
        musescore_paths = [
            r"C:\\Program Files\\MuseScore 4\\bin\\MuseScore4.exe",
            r"C:\\Program Files (x86)\\MuseScore 4\\bin\\MuseScore4.exe",
            r"C:\\Program Files\\MuseScore 3\\bin\\MuseScore3.exe",
            r"C:\\Program Files (x86)\\MuseScore 3\\bin\\MuseScore3.exe",
        ]

    for path in musescore_paths:
        if os.path.exists(path):
            return path

    return (
        shutil.which("mscore")
        or shutil.which("MuseScore4")
        or shutil.which("MuseScore3")
        or shutil.which("MuseScore4.exe")
        or shutil.which("MuseScore3.exe")
    )


def extract_midi_from_mscx(mscx_file_path, output_file_path=None, measure_range=None):
    """Extract MIDI from .mscx or .mscz file using MuseScore CLI or mido library

    Args:
        mscx_file_path: Path to the MuseScore file
        output_file_path: Path for output MIDI file (optional, auto-generated if None)
        measure_range: Tuple (start_measure, end_measure) to extract only specific measures (1-indexed, inclusive)
                       If None, extracts all measures. Note: Only supported with library-based extraction.

    Returns:
        Path to the created MIDI file, or None if extraction failed
    """
    output_dir = os.path.normpath(str(MIDI_OUTPUT_DIR))
    os.makedirs(output_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(mscx_file_path))[0]
    if output_file_path is None:
        filename = base_name + ".mid"
        output_file_path = os.path.normpath(os.path.join(output_dir, filename))

    musescore_exe = _find_musescore_exe()

    if musescore_exe and measure_range is None:
        try:
            result = subprocess.run(
                [musescore_exe, "-o", output_file_path, mscx_file_path],
                capture_output=True,
                text=True,
                timeout=30,
            )
            if result.returncode == 0 and os.path.exists(output_file_path):
                return output_file_path

            result = subprocess.run(
                [musescore_exe, mscx_file_path, "-o", output_file_path],
                capture_output=True,
            )
            if result.returncode == 0 and os.path.exists(output_file_path):
                return output_file_path
        except (subprocess.TimeoutExpired, FileNotFoundError, Exception):
            pass

    try:
        import mido
        from mido import Message, MidiFile, MidiTrack

        if mscx_file_path.endswith(".mscz"):
            with zipfile.ZipFile(mscx_file_path, "r") as zip_ref:
                file_list = zip_ref.namelist()
                score_file = None
                for name in file_list:
                    if name.endswith(".mscx") or ("." not in name and not name.endswith("/")):
                        score_file = name
                        break
                if score_file is None:
                    score_file = file_list[0]
                with zip_ref.open(score_file) as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
        else:
            tree = ET.parse(mscx_file_path)
            root = tree.getroot()

        division = 480
        for elem in root.iter("Division"):
            if elem.text:
                division = int(elem.text)
                break

        mid = MidiFile()
        track = MidiTrack()
        mid.tracks.append(track)

        tempo = mido.bpm2tempo(DEFAULT_MIDI_TEMPO)
        track.append(Message("set_tempo", tempo=tempo, time=0))

        notes = []

        all_measures = list(root.iter("Measure"))
        measures_to_process = []
        start_tick_offset = 0

        if measure_range:
            start_measure, end_measure = measure_range
            for idx, measure in enumerate(all_measures):
                measure_no_attr = measure.get("no")
                if measure_no_attr is not None:
                    try:
                        measure_no = int(measure_no_attr)
                    except (ValueError, TypeError):
                        measure_no = idx + 1
                else:
                    measure_no = idx + 1

                if measure_no < start_measure:
                    start_tick_offset += division * 4
                elif start_measure <= measure_no <= end_measure:
                    measures_to_process.append((measure_no, measure))
        else:
            measures_to_process = [(idx + 1, m) for idx, m in enumerate(all_measures)]

        current_tick = 0
        for measure_no, measure in measures_to_process:
            measure_tick = 0

            for chord in measure.iter("Chord"):
                chord_tick = measure_tick
                if "tick" in chord.attrib:
                    tick_val = int(chord.attrib["tick"])
                    if tick_val < division * 4:
                        chord_tick = tick_val
                    else:
                        chord_tick = tick_val - start_tick_offset

                duration = division
                duration_elem = chord.find("duration")
                if duration_elem is not None and duration_elem.text:
                    duration = int(duration_elem.text)

                for note in chord.findall("Note"):
                    pitch_elem = note.find("pitch")
                    if pitch_elem is not None and pitch_elem.text:
                        midi_pitch = int(pitch_elem.text)
                        notes.append((current_tick + chord_tick, midi_pitch, duration))

                measure_tick = max(measure_tick, chord_tick + duration)

            current_tick += measure_tick if measure_tick > 0 else division * 4

        notes.sort(key=lambda x: x[0])

        last_tick = 0
        for tick, pitch, duration in notes:
            delta = tick - last_tick
            track.append(Message("note_on", channel=0, note=pitch, velocity=64, time=delta))
            track.append(Message("note_off", channel=0, note=pitch, velocity=64, time=duration))
            last_tick = tick + duration

        mid.save(output_file_path)
        return output_file_path

    except ImportError:
        try:
            from music21 import converter

            score = converter.parse(mscx_file_path)
            score.write("midi", output_file_path)
            return output_file_path
        except ImportError:
            raise ImportError(
                "MIDI extraction requires either:\n"
                "1. MuseScore command-line tool installed, or\n"
                "2. Python library 'mido' (pip install mido), or\n"
                "3. Python library 'music21' (pip install music21)"
            )
    except Exception as e:
        raise Exception(f"Failed to extract MIDI: {str(e)}")
