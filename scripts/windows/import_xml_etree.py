import xml.etree.ElementTree as ET
import os
import sys
from pathlib import Path

def get_pitch_name(pitch_value):
    """Convert MIDI pitch number to note name (e.g., 60 -> C4)"""
    pitch_names = ["C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B"]
    note = pitch_value % 12
    octave = (pitch_value // 12) - 1
    return f"{pitch_names[note]}{octave}"

def extract_pitches_from_mscx(mscx_file_path, output_file_path=None):
    """Extract pitch names from a MuseScore .mscx file"""
    try:
        tree = ET.parse(mscx_file_path)
        root = tree.getroot()
        
        # MuseScore files use MusicXML namespace
        ns = {'m': 'http://www.musescore.org/mscx'}
        
        pitches = []
        
        # Find all note elements
        for note in root.findall('.//m:Note', ns):
            # Get pitch information
            pitch_elem = note.find('m:pitch', ns)
            if pitch_elem is not None:
                step = pitch_elem.find('m:step', ns)
                octave = pitch_elem.find('m:octave', ns)
                alter = pitch_elem.find('m:alter', ns)
                
                if step is not None and octave is not None:
                    step_text = step.text
                    octave_num = int(octave.text)
                    
                    # Handle accidentals
                    if alter is not None:
                        alter_val = int(alter.text)
                        if alter_val == 1:
                            step_text += "#"
                        elif alter_val == -1:
                            step_text += "b"
                    
                    pitch_name = f"{step_text}{octave_num}"
                    pitches.append(pitch_name)
            else:
                # Try to get MIDI pitch if available
                midi_pitch = note.find('m:pitch', ns)
                if midi_pitch is None:
                    # Check for chord/rest
                    rest = note.find('m:rest', ns)
                    if rest is None:
                        # Try to calculate from tick or other methods
                        pass
        
        # If no pitches found with MusicXML namespace, try without namespace
        if not pitches:
            for note in root.findall('.//Note'):
                pitch_elem = note.find('pitch')
                if pitch_elem is not None:
                    step = pitch_elem.find('step')
                    octave = pitch_elem.find('octave')
                    alter = pitch_elem.find('alter')
                    
                    if step is not None and octave is not None:
                        step_text = step.text
                        octave_num = int(octave.text)
                        
                        if alter is not None:
                            alter_val = int(alter.text)
                            if alter_val == 1:
                                step_text += "#"
                            elif alter_val == -1:
                                step_text += "b"
                        
                        pitch_name = f"{step_text}{octave_num}"
                        pitches.append(pitch_name)
        
        # Write to output file
        if output_file_path is None:
            output_file_path = os.path.splitext(mscx_file_path)[0] + "_pitches.txt"
        
        with open(output_file_path, 'w', encoding='utf-8') as f:
            for pitch in pitches:
                f.write(pitch + '\n')
        
        print(f"Extracted {len(pitches)} pitches to: {output_file_path}")
        return pitches
        
    except Exception as e:
        print(f"Error processing file: {e}")
        return None

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python -m music_clipboard.extract.pitches <path_to_mscx_file> [output_file]")
        print("\nExample:")
        print("  python -m music_clipboard.extract.pitches score.mscx")
        print("  python -m music_clipboard.extract.pitches score.mscx output.txt")
        sys.exit(1)
    
    mscx_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    if not os.path.exists(mscx_file):
        print(f"Error: File '{mscx_file}' not found.")
        sys.exit(1)
    
    extract_pitches_from_mscx(mscx_file, output_file)
