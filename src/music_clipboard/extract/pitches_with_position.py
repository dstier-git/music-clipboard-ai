import os
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

if __package__ is None or __package__ == "":
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from music_clipboard.platform.runtime import output_dirs

OUTPUT_DIR = output_dirs()[0]


def get_pitch_name(pitch_value):
    """Convert MIDI pitch number to note name (e.g., 60 -> C4)."""
    pitch_names = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"]
    note = pitch_value % 12
    octave = (pitch_value // 12) - 1
    return f"{pitch_names[note]}{octave}"


def get_division(root):
    """Get the Division value from the score (ticks per quarter note)."""
    division = 480
    for elem in root.iter("Division"):
        if elem.text:
            division = int(elem.text)
            break
    return division


def get_time_signature(measure):
    """Get time signature from a measure element."""
    for ts in measure.iter("TimeSig"):
        sig_n = ts.find("sigN")
        sig_d = ts.find("sigD")
        if sig_n is not None and sig_d is not None:
            return int(sig_n.text), int(sig_d.text)
    return None, None


def tick_to_measure_beat(tick, measures, division):
    """Convert absolute tick to measure number and beat position."""
    current_tick = 0
    measure_num = 1

    for measure in measures:
        measure_length = 0
        time_sig_n, time_sig_d = get_time_signature(measure)

        if time_sig_n and time_sig_d:
            measure_length = (time_sig_n * division * 4) // time_sig_d
        else:
            measure_length = division * 4

        if current_tick + measure_length > tick:
            beat_tick = tick - current_tick
            beat = (beat_tick / division) + 1
            return measure_num, beat

        current_tick += measure_length
        measure_num += 1

    measure_num = (tick // (division * 4)) + 1
    beat = (tick % (division * 4)) / division + 1
    return measure_num, beat


def extract_pitches_with_position_from_mscx(mscx_file_path, output_file_path=None, debug=True, measure_range=None):
    """Extract pitch names and metric positions from a MuseScore .mscx or .mscz file.

    Args:
        mscx_file_path: Path to the MuseScore file
        output_file_path: Path for output file (optional)
        debug: Whether to print debug information
        measure_range: Tuple (start_measure, end_measure) to extract only specific measures (1-indexed, inclusive)
                       If None, extracts all measures
    """
    try:
        if mscx_file_path.endswith(".mscz"):
            print("Detected .mscz file, extracting...")
            with zipfile.ZipFile(mscx_file_path, "r") as zip_ref:
                file_list = zip_ref.namelist()
                score_file = None
                for name in file_list:
                    if name.endswith(".mscx") or ("." not in name and not name.endswith("/")):
                        score_file = name
                        break

                if score_file is None:
                    score_file = file_list[0]

                print(f"Reading {score_file} from archive...")
                with zip_ref.open(score_file) as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
        else:
            tree = ET.parse(mscx_file_path)
            root = tree.getroot()

        division = get_division(root)
        if debug:
            print(f"Division (ticks per quarter note): {division}")
            print(f"\nRoot tag: {root.tag}")
            print(f"Root attributes: {root.attrib}")

        all_measures = list(root.iter("Measure"))
        if debug:
            print(f"Found {len(all_measures)} measures")

        measures_with_numbers = []
        for idx, measure in enumerate(all_measures):
            measure_no_attr = measure.get("no")
            if measure_no_attr is not None:
                try:
                    measure_no = int(measure_no_attr)
                except (ValueError, TypeError):
                    measure_no = idx + 1
            else:
                measure_no = idx + 1
            measures_with_numbers.append((measure_no, measure))

        if measure_range:
            start_measure, end_measure = measure_range
            measures_with_numbers = [
                (no, m) for no, m in measures_with_numbers if start_measure <= no <= end_measure
            ]
            if debug:
                print(
                    f"Filtering to measures {start_measure}-{end_measure}: {len(measures_with_numbers)} measures found"
                )

        notes_with_position = []
        current_tick = 0

        if measure_range:
            start_measure, end_measure = measure_range
            for idx, measure in enumerate(all_measures):
                measure_no_attr = measure.get("no")
                if measure_no_attr is not None:
                    try:
                        actual_no = int(measure_no_attr)
                    except (ValueError, TypeError):
                        actual_no = idx + 1
                else:
                    actual_no = idx + 1

                if actual_no < start_measure:
                    time_sig_n, time_sig_d = get_time_signature(measure)
                    if time_sig_n and time_sig_d:
                        measure_length = (time_sig_n * division * 4) // time_sig_d
                    else:
                        measure_length = division * 4
                    current_tick += measure_length
                else:
                    break

        for current_measure_num, measure in measures_with_numbers:
            time_sig_n, time_sig_d = get_time_signature(measure)
            if time_sig_n and time_sig_d:
                measure_length = (time_sig_n * division * 4) // time_sig_d
            else:
                measure_length = division * 4

            measure_tick = 0
            has_chords = False

            for chord in measure.iter("Chord"):
                has_chords = True
                chord_tick = None

                if "tick" in chord.attrib:
                    chord_tick_val = int(chord.attrib["tick"])
                    if chord_tick_val < measure_length:
                        chord_tick = current_tick + chord_tick_val
                    else:
                        chord_tick = chord_tick_val
                else:
                    chord_tick = current_tick + measure_tick

                duration_elem = chord.find("duration")
                if duration_elem is not None and duration_elem.text:
                    duration = int(duration_elem.text)
                    measure_tick = max(measure_tick, chord_tick - current_tick if chord_tick else measure_tick)
                else:
                    duration = division

                for note in chord.findall("Note"):
                    pitch_elem = note.find("pitch")
                    if pitch_elem is not None and pitch_elem.text:
                        midi_pitch = int(pitch_elem.text)
                        pitch_name = get_pitch_name(midi_pitch)

                        if chord_tick is not None:
                            beat_tick = chord_tick - current_tick
                            beat = (beat_tick / division) + 1
                            position_str = f"M{current_measure_num}:{beat:.2f}"
                            notes_with_position.append((pitch_name, position_str, chord_tick))
                        else:
                            beat = (measure_tick / division) + 1
                            position_str = f"M{current_measure_num}:{beat:.2f}"
                            notes_with_position.append(
                                (pitch_name, position_str, current_tick + measure_tick)
                            )

                        if debug and len(notes_with_position) <= 5:
                            tick_val = chord_tick or (current_tick + measure_tick)
                            print(f"Found note: {pitch_name} at {position_str} (tick: {tick_val})")

                if duration_elem is not None and duration_elem.text:
                    measure_tick += int(duration_elem.text)
                else:
                    measure_tick += division

            if not has_chords:
                for note in measure.iter("Note"):
                    pitch_elem = note.find("pitch")
                    if pitch_elem is not None and pitch_elem.text:
                        midi_pitch = int(pitch_elem.text)
                        pitch_name = get_pitch_name(midi_pitch)

                        beat = (measure_tick / division) + 1
                        position_str = f"M{current_measure_num}:{beat:.2f}"

                        notes_with_position.append((pitch_name, position_str, current_tick + measure_tick))
                        if debug and len(notes_with_position) <= 5:
                            print(f"Found note: {pitch_name} at {position_str}")

                        measure_tick += division

            current_tick += measure_length

        if not notes_with_position:
            print("Trying fallback approach...")
            for chord in root.iter("Chord"):
                for note in chord.findall("Note"):
                    pitch_elem = note.find("pitch")
                    if pitch_elem is not None and pitch_elem.text:
                        midi_pitch = int(pitch_elem.text)
                        pitch_name = get_pitch_name(midi_pitch)
                        notes_with_position.append((pitch_name, "M?:?", None))
                        if debug and len(notes_with_position) <= 5:
                            print(f"Found note (fallback): {pitch_name}")

        if not notes_with_position:
            ns = {"m": "http://www.musescore.org/mscx"}
            for note in root.findall(".//m:Note", ns):
                pitch_elem = note.find("m:pitch", ns)
                if pitch_elem is not None and pitch_elem.text:
                    midi_pitch = int(pitch_elem.text)
                    pitch_name = get_pitch_name(midi_pitch)
                    notes_with_position.append((pitch_name, "M?:?", None))

        if debug:
            print(f"\n{'=' * 50}")
            if notes_with_position:
                print(f"Successfully extracted {len(notes_with_position)} notes with positions!")
            else:
                print("No notes found. Checking for Note elements...")
                note_count = len(list(root.iter("Note")))
                chord_count = len(list(root.iter("Chord")))
                print(f"Found {note_count} Note elements")
                print(f"Found {chord_count} Chord elements")

        if notes_with_position:
            output_dir = os.path.normpath(str(OUTPUT_DIR))
            os.makedirs(output_dir, exist_ok=True)

            base_name = os.path.splitext(os.path.basename(mscx_file_path))[0]
            filename = base_name + "_pitches_with_position.txt"
            output_file_path = os.path.normpath(os.path.join(output_dir, filename))

            if debug:
                print(f"DEBUG: Saving to OUTPUT_DIR: {output_dir}")
                print(f"DEBUG: Filename: {filename}")
                print(f"DEBUG: Full output path: {output_file_path}")

            with open(output_file_path, "w", encoding="utf-8") as f:
                for pitch, position, tick in notes_with_position:
                    if tick is not None:
                        f.write(f"{pitch}\t{position}\t(tick: {tick})\n")
                    else:
                        f.write(f"{pitch}\t{position}\n")

            print(f"Extracted {len(notes_with_position)} notes to: {output_file_path}")
            return notes_with_position, output_file_path

        return notes_with_position, None

    except Exception as e:
        print(f"Error processing file: {e}")
        import traceback
        traceback.print_exc()
        return None


def main():
    """Interactive main function."""
    print("=" * 60)
    print("MuseScore Pitch & Position Extractor")
    print("=" * 60)
    print()

    file_path = input("Enter the path to your MuseScore file (.mscx or .mscz): ").strip()
    file_path = file_path.strip('"').strip("'")

    if not os.path.exists(file_path):
        print(f"\nError: File not found: {file_path}")
        return

    if not (file_path.endswith(".mscx") or file_path.endswith(".mscz")):
        print("\nWarning: File doesn't have .mscx or .mscz extension. Continuing anyway...")

    print(f"\nProcessing: {file_path}")
    print("-" * 60)

    result = extract_pitches_with_position_from_mscx(file_path, debug=True)

    if isinstance(result, tuple) and len(result) == 2:
        notes, actual_output_path = result
    else:
        notes = result
        actual_output_path = None

    if notes:
        print(f"\n{'=' * 60}")
        print("Extraction complete!")
        print(f"Total notes extracted: {len(notes)}")
        print("\nFirst 20 notes (Pitch | Position):")
        for i, (pitch, position, tick) in enumerate(notes[:20], 1):
            print(f"  {i}. {pitch} | {position}")
        if len(notes) > 20:
            print(f"  ... and {len(notes) - 20} more")

        if actual_output_path:
            print(f"\nNotes saved to: {actual_output_path}")
        else:
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            output_file = os.path.join(str(OUTPUT_DIR), base_name + "_pitches_with_position.txt")
            print(f"\nNotes saved to: {output_file}")
        print("Format: Pitch | Measure:Beat | (tick position)")
    else:
        print("\nNo notes extracted. Please check the debug output above.")


if __name__ == "__main__":
    main()
