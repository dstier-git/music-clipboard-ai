from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[2]
SRC_ROOT = PROJECT_ROOT / "src"
for root in (PROJECT_ROOT, SRC_ROOT):
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))

from music_clipboard.gui.app import main


if __name__ == "__main__":
    print(
        "Deprecated legacy entrypoint: clipboard-full/WIN/musescore_extractor_gui.py. "
        "Use: python -m music_clipboard.gui.app",
        file=sys.stderr,
    )
    main()
