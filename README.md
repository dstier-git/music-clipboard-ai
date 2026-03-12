# Music Clipboard

Music Clipboard with (optional) AI integration helps you capture, transfer, and edit music regions within or across DAWs/notation software. The program supports extraction optimized for cross-DAW transfer and translates music to text for language models.

Music Clipboard was built with both technical and non-technical musicians in mind, after the setup below the GUI requires *no coding or terminal usage.*

Currently up for **macOS**, Windows version is being updated.

## Quickstart

Run these commands from the repository root (`music_clipboard_ai/`).

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Minimal Setup for `electron_ui`

If you only want to run the Electron UI, from the repository root:

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
cd src/music_clipboard/electron_ui/app
npm install
npm start
```

## Electron + FastAPI App (New)

The new desktop app lives at:

- `src/music_clipboard/electron_ui/app` (Electron Forge + React + TypeScript)
- `src/music_clipboard/electron_ui/backend` (FastAPI local backend)

### Run the desktop app in development

From repository root:

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
cd src/music_clipboard/electron_ui/app
npm install
npm start
```

The Electron main process will start the local FastAPI backend automatically on `127.0.0.1` with an ephemeral port and token.

### Package the desktop app

From `src/music_clipboard/electron_ui/app`:

```bash
source ../../../../venv/bin/activate
pip install pyinstaller
npm run bundle:backend
npm run package
```

`bundle:backend` produces a bundled backend executable in `backend-dist/`, and Forge packages it as an app resource.

### Backend API (v1)

- `GET /api/v1/health`
- `POST /api/v1/extract`
- `POST /api/v1/midi/export`
- `POST /api/v1/ai/openai/edit-midi`
- `GET /api/v1/jobs/{job_id}`
- `GET /api/v1/jobs/{job_id}/events` (SSE)

Launch the GUI:

```bash
scripts/macos/run_gui.sh
```

Or launch directly with Python (cross-platform):

```bash
PYTHONPATH="$(pwd)/src:$(pwd)" python3 -m music_clipboard.gui.app
```

Start the global hotkey listener:

```bash
scripts/macos/run_hotkey_listener.sh
```

Or launch directly with Python:

```bash
PYTHONPATH="$(pwd)/src:$(pwd)" python3 -m music_clipboard.automation.hotkey_listener
```

## Platform Commands

### macOS

GUI launcher:

```bash
scripts/macos/run_gui.sh
```

Hotkey listener:

```bash
scripts/macos/run_hotkey_listener.sh
```

### Windows

GUI launcher:

```bash
scripts\\windows\\run_gui.bat
```

Hotkey listener:

```bash
scripts\\windows\\run_hotkey_listener.bat
```

Python module alternatives (PowerShell):

```bash
$env:PYTHONPATH = "$PWD\\src;$PWD"
python -m music_clipboard.gui.app
python -m music_clipboard.automation.hotkey_listener
```

## Outputs

Generated files are written to:

- `data/outputs/text` for extracted pitch text files
- `data/outputs/midi` for exported MIDI files

During transition, legacy output folders may still be read if present:

- `clipboard-full/txts`
- `clipboard-full/midis`

## Feature Reference

- **AI-integrated workflow:** copy/export extracted musical material into AI editing flows from the GUI
    - Supports editing within cross-DAW transfer
- Hotkey automation: global shortcut listener to trigger save-selection workflows
- Inputs: MuseScore `.mscx` and `.mscz` scores (plus MIDI handling in app workflows)
- Pitch extraction: note names from score content (for example `C4`, `F#5`)
- Position-aware extraction: note names with measure/beat locations (for example `M12:3.50`)
- MIDI export: MIDI generation from MuseScore scores

## MuseScore MCP Integration

This repository also includes a MuseScore MCP integration under:

- `src/music_clipboard/integrations/musescore_mcp`

Entry points:

- `src/music_clipboard/integrations/musescore_mcp/mcp_server.py`
- `src/music_clipboard/integrations/musescore_mcp/musescore_mcp_websocket.qml`

For setup and usage, see:

- `src/music_clipboard/integrations/musescore_mcp/README.md`

## Legacy Compatibility

`clipboard-full/MAC` and `clipboard-full/WIN` are legacy transitional wrappers. They remain in the repo for compatibility, but new usage should prefer `src/music_clipboard` modules and `scripts/macos` / `scripts/windows` launchers.

## Acknowledgements

The MuseScore MCP integration in this codebase was originally based on [ghchen99/mcp-musescore](https://github.com/ghchen99/mcp-musescore) and then integrated and modified in this repository.
