from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[2]
SRC_ROOT = PROJECT_ROOT / "src"
for root in (PROJECT_ROOT, SRC_ROOT):
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))

from music_clipboard.extract.pitches import extract_pitches_from_mscx, main

__all__ = ["extract_pitches_from_mscx", "main"]

if __name__ == "__main__":
    print(
        "Deprecated legacy entrypoint: clipboard-full/WIN/extract_pitches.py. "
        "Use: python -m music_clipboard.extract.pitches",
        file=sys.stderr,
    )
    main()
