import os
import sys
import traceback
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

if __package__ is None or __package__ == "":
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from music_clipboard.platform.runtime import output_dirs
from music_clipboard.extract.midi import extract_midi_from_mscx

OUTPUT_DIR = output_dirs()[0]


def get_pitch_name(pitch_value):
    """Convert MIDI pitch number to note name (e.g., 60 -> C4)."""
    pitch_names = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"]
    note = pitch_value % 12
    octave = (pitch_value // 12) - 1
    return f"{pitch_names[note]}{octave}"


def extract_pitches_from_mscx(mscx_file_path, output_file_path=None, debug=False):
    """Extract pitch names from a MuseScore .mscx or .mscz file."""
    try:
        if mscx_file_path.endswith(".mscz"):
            if debug:
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
                if debug:
                    print(f"Reading {score_file} from archive...")
                with zip_ref.open(score_file) as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
        else:
            tree = ET.parse(mscx_file_path)
            root = tree.getroot()

        if debug:
            print(f"\nRoot tag: {root.tag}")
            print(f"Root attributes: {root.attrib}")
            print("\nXML Structure (first 20 elements):")
            count = 0
            for elem in root.iter():
                if count < 20:
                    sample = elem.text[:50] if elem.text and elem.text.strip() else ""
                    print(f"  {elem.tag}: {sample}")
                    count += 1

        pitches = []

        for chord in root.iter("Chord"):
            for note in chord.findall("Note"):
                pitch_elem = note.find("pitch")
                if pitch_elem is not None and pitch_elem.text:
                    midi_pitch = int(pitch_elem.text)
                    pitch_name = get_pitch_name(midi_pitch)
                    pitches.append(pitch_name)
                    if debug and len(pitches) <= 5:
                        print(f"Found note (Chord/Note/pitch): {pitch_name} (MIDI: {midi_pitch})")

        if not pitches:
            for note in root.iter("Note"):
                pitch_elem = note.find("pitch")
                if pitch_elem is not None and pitch_elem.text:
                    midi_pitch = int(pitch_elem.text)
                    pitch_name = get_pitch_name(midi_pitch)
                    pitches.append(pitch_name)
                    if debug and len(pitches) <= 5:
                        print(f"Found note (Note/pitch): {pitch_name} (MIDI: {midi_pitch})")

        if not pitches:
            ns = {"m": "http://www.musescore.org/mscx"}
            for note in root.findall(".//m:Note", ns):
                pitch_elem = note.find("m:pitch", ns)
                if pitch_elem is not None and pitch_elem.text:
                    midi_pitch = int(pitch_elem.text)
                    pitch_name = get_pitch_name(midi_pitch)
                    pitches.append(pitch_name)

        if debug:
            print(f"\n{'=' * 50}")
            if pitches:
                print(f"Successfully extracted {len(pitches)} pitches!")
            else:
                print("No pitches found. Checking for Note elements...")
                note_count = len(list(root.iter("Note")))
                chord_count = len(list(root.iter("Chord")))
                print(f"Found {note_count} Note elements")
                print(f"Found {chord_count} Chord elements")

        if pitches:
            output_dir = os.path.normpath(str(OUTPUT_DIR))
            os.makedirs(output_dir, exist_ok=True)

            base_name = os.path.splitext(os.path.basename(mscx_file_path))[0]
            filename = base_name + "_pitches.txt"
            output_file_path = os.path.normpath(os.path.join(output_dir, filename))

            with open(output_file_path, "w", encoding="utf-8") as f:
                for pitch in pitches:
                    f.write(pitch + "\n")

            print(f"Extracted {len(pitches)} pitches to: {output_file_path}")

        return pitches

    except Exception as e:
        print(f"Error processing file: {e}")
        traceback.print_exc()
        return None


def main():
    """Interactive main function."""
    print("-" * 60)
    print("MuseScore Pitch Extractor")
    print("-" * 60)
    print()

    file_path = input("Enter the path to MuseScore file (.mscx or .mscz): ").strip()
    file_path = file_path.strip('"').strip("'")

    if not os.path.exists(file_path):
        print(f"\nError: File not found: {file_path}")
        return

    if not (file_path.endswith(".mscx") or file_path.endswith(".mscz")):
        print("\nWarning: File doesn't have .mscx or .mscz extension...")

    print("\nOutput format:")
    print("1. Text (pitch names)")
    print("2. MIDI")
    format_choice = input("Select format (1 or 2, default: 1): ").strip() or "1"

    print(f"\nProcessing: {file_path}\n")

    try:
        if format_choice == "2":
            midi_path = extract_midi_from_mscx(file_path)
            if midi_path:
                print(f"\n{'-' * 60}")
                print("MIDI extraction complete!")
                print(f"MIDI file saved to: {midi_path}")
            else:
                print("\nFailed to extract MIDI.")
        else:
            pitches = extract_pitches_from_mscx(file_path, debug=True)
            if pitches:
                print(f"\n{'-' * 60}")
                print("Extraction complete!")
                print(f"Total notes extracted: {len(pitches)}")
                print("\nFirst 20 pitches:")
                for i, pitch in enumerate(pitches[:20], 1):
                    print(f"  {i}. {pitch}")
                if len(pitches) > 20:
                    print(f"  ... and {len(pitches) - 20} more")

                base_name = os.path.splitext(os.path.basename(file_path))[0]
                output_file = os.path.join(str(OUTPUT_DIR), base_name + "_pitches.txt")
                print(f"\nNotes saved to: {output_file}")
            else:
                print("\nNo pitches extracted.")
    except Exception as e:
        print(f"\nError processing file: {e}")
        traceback.print_exc()


if __name__ == "__main__":
    main()
