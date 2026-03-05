# Music Clipboard with AI

Cross-platform tooling for extracting pitch names and metric positions from MuseScore files (`.mscx`/`.mscz`) or MIDI data, then handing results into AI-assisted music editing workflows.

## Canonical layout

```text
music_clipboard_ai/
  src/music_clipboard/
    gui/app.py
    extract/{midi.py,pitches.py,pitches_with_position.py}
    automation/hotkey_listener.py
    platform/runtime.py
    integrations/musescore_mcp/
  scripts/
    macos/{run_gui.sh,run_hotkey_listener.sh}
    windows/{run_gui.bat,run_hotkey_listener.bat,run_gui.sh}
  data/outputs/{text,midi}
  requirements.txt
```

## Transitional compatibility (Phase 1)

Legacy launchers and wrappers under `clipboard-full/MAC` and `clipboard-full/WIN` are still present.
They print deprecation warnings and forward to the new module paths.

## Run

Preferred:

```bash
python -m music_clipboard.gui.app
```

Or use canonical launchers:

```bash
scripts/macos/run_gui.sh
scripts/windows/run_gui.bat
```

Hotkey listener:

```bash
python -m music_clipboard.automation.hotkey_listener
```

## Output locations

- New write targets:
  - `data/outputs/text`
  - `data/outputs/midi`
- Legacy fallback directories still recognized during transition:
  - `clipboard-full/txts`
  - `clipboard-full/midis`

## Dependencies

Install from root:

```bash
pip install -r requirements.txt
```

## MuseScore MCP integration

Location:

- `src/music_clipboard/integrations/musescore_mcp`

Entry points:

- `mcp_server.py`
- `musescore_mcp_websocket.qml`

See integration docs at:

- `src/music_clipboard/integrations/musescore_mcp/README.md`

## Acknowledgements

The MuseScore MCP integration in this codebase was originally based on [ghchen99/mcp-musescore](https://github.com/ghchen99/mcp-musescore) and integrated/modified in this repository.
