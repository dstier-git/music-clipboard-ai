# Music Clipboard with AI

A cross-platform music clipboard for extracting pitch names and metric positions from MuseScore files (`.mscx`/`.mscz`) or MIDI data and handing them off to AI-assisted editing workflows. The GUI and scripts share one unified implementation with OS-specific behavior handled automatically, so you can use manual or automated flows on macOS and Windows without maintaining separate codebases.

The core idea: capture the exact musical fragment you want, then let AI help turn rough joins into transitions that feel smoother and more refined.

## Repository layout
- `app/` - Unified, cross-platform backend infrastructure.
- `MAC/` - macOS wrappers and shell helpers (e.g., `run_gui.sh`).
- `WIN/` - Windows wrappers, batch scripts, and helpers.
- `midis/`, `txts/`, and other shared folders store extracted clipboard artifacts created by either platform.

## Requirements
1. Python 3.6 or newer.
2. Standard, shipped Python libraries:
   - `tkinter` for the GUI.
   - `xml.etree.ElementTree` for parsing MuseScore XML.
   - `zipfile` to read compressed `.mscz` scores.
3. Optional automation dependencies (install with `pip install -r requirements.txt` or individually):
   ```bash
   pip install pyautogui pywinauto keyboard psutil
   ```
   - `pyautogui` is required for the automation helpers (triggering Save Selection) on both platforms.
   - `pywinauto` and `keyboard` power the Windows automation and global hotkey listener.
   - `psutil` helps the worker scripts detect MuseScore and coordinate with the hotkey listener.

> The GUI runs without optional deps, but automation buttons, hotkey listener, and background helpers stay disabled.

## Platform-specific workflows

### macOS (`MAC/`)

1. Launch the GUI with:
   ```bash
   python3 musescore_extractor_gui.py
   ```
   or via `chmod +x run_gui.sh && ./run_gui.sh`.
2. **Manual Mode**: Browse for a MuseScore file, then click **Extract**; output appears in the GUI and is saved under `txts/` in the repository root.
3. **Auto Mode**:
   - Set the watch folder (default `Documents/MuseScore4/Scores`) and click **Start Watching**.
   - For each accepted new `.mscx`/`.mscz` file, the app now:
     - Runs the extraction workflow.
     - Prompts for an AI edit instruction.
     - Sends `Connect to musescore and ...` to Claude Desktop (auto-paste + Enter), then lets Claude MCP execute.
     - Enables quick AI polishing prompts focused on smoother, cleaner transitions between phrases.
   - MIDI watch support: new `.mid`/`.midi` files are also accepted and opened in MuseScore for the AI flow.
     - MIDI inputs skip extraction and only run prompt -> MuseScore open -> Claude send.
   - In MuseScore, select measures and:
     - Click **Trigger Save Selection in MuseScore** in the app (requires `pyautogui`).
     - Or save selection manually via **File > Save Selection** (`Shift+Cmd+S`).
   - MuseScore saves the selection to the watched folder and the app detects it automatically as a clipboard-ready fragment.
   - Rate limit: only one new score file is processed per 60-second window. Additional score files in that window are ignored with no action.
   - Claude requirement: Claude Desktop must already be running before MuseScore is opened for AI automation.
   - If opening MuseScore 4 fails, the app shows an error and cancels Claude sending for that file.
4. macOS automation relies on AppleScript; grant accessibility permissions to Terminal/Python under **System Preferences > Security & Privacy > Privacy > Accessibility** if automation buttons stay disabled.

### Windows (`WIN/`)

1. Run the GUI via:
   ```bash
   python musescore_extractor_gui.py
   ```
2. **Manual Mode**: Browse for a score, optionally enter a measure range, then click **Extract**; output files are saved next to the input file by default.
3. **Auto Mode**:
   - Choose a watch folder (default `Documents/MuseScore4/Scores`) and start watching.
   - In MuseScore, select measures and:
     - Hit the global hotkey `Ctrl+Alt+S` (after starting `hotkey_listener.py` or `run_hotkey_listener.bat`) to trigger Save Selection even when MuseScore runs in the background.
     - Or click **Trigger Save Selection** inside the GUI.
     - Or use MuseScore's native **File > Save Selection** (`Ctrl+Shift+S`).
   - Saved selections go into the watched folder and the app processes them automatically for clipboard + AI refinement flow.
4. To always listen for hotkeys, run `python hotkey_listener.py` or `run_hotkey_listener.bat` (add the batch file to `%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup` for persistence). The listener writes requests to `musescore_hotkey_request.txt`, so the GUI can react even when it launches after a keypress.

## Shared command-line scripts

The scripts available in each platform folder behave the same and are thin wrappers around the unified `app/` implementation:

```bash
python extract_pitches.py
python extract_pitches_with_position.py
```

Both prompt for a MuseScore file path, extract pitch names and optional metric positions, and produce a simple text file in the target folder. These text outputs are ideal for quick AI prompt context when refining musical continuity.

## Output format
- **Pitch**: Note name (e.g., `C4`, `E5`, `F#3`).
- **Position**: Measure and beat (e.g., `M1:1.00`, `M2:2.50`).
- **Tick**: Internal timing number for reference.

Example:
```
C4	M1:1.00	(tick: 0)
E4	M1:1.00	(tick: 0)
G4	M1:2.00	(tick: 480)
```

## Tips
- Use MuseScore's **Save Selection** to extract only the measures you care about.
- Auto mode is ideal for repeated clipboard captures; manual mode works well for one-off full scores.
- On macOS, the GUI previews the first 10 notes in the output region.
- You can clear the output area at any time if it becomes crowded.
- Saved selections already restrict the score to the selected measures, so no manual range is needed.
- For better AI edits, include intent in your instruction (for example: "make the transition into the chorus smoother, keep harmonic tension, and avoid abrupt leaps").

## Troubleshooting
- **Import errors**: Ensure the `app/` directory is present and the platform wrapper scripts (`MAC/` or `WIN/`) have not been moved out of the repository.
- **No notes extracted**: Confirm the file is a valid `.mscx`/`.mscz` MuseScore file.
- **Watch folder not detecting files**: Verify MuseScore saves selections to the configured folder.
- **A detected file appears to be ignored**: In macOS AI auto mode, only one file is accepted every 60 seconds; later files in that window are ignored silently.
- **MIDI file behavior**: `.mid`/`.midi` files are opened in MuseScore for AI flow, but extraction is skipped for MIDI inputs.
- **"Claude Not Running"**: Open Claude Desktop first; the app will not auto-launch Claude.
- **"MuseScore Open Failed"**: Ensure MuseScore 4 is installed as `MuseScore 4` and can be opened manually. Claude sending is canceled when this occurs.
- **Automation buttons stay disabled**: Install the optional dependencies (`pyautogui`, `pywinauto`, `keyboard`) and grant accessibility permissions (macOS).
- **"MuseScore Not Found"**: Start MuseScore 4 with at least one score open before triggering automation.
- **Global hotkey issues** (Windows): Run the listener helper as Administrator if necessary, and ensure the `keyboard` library is installed.
- **AppleScript permissions** (macOS): Add Terminal or your Python IDE to the Accessibility list under **Security & Privacy**.
