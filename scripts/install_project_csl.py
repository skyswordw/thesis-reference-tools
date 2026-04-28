from __future__ import annotations

import argparse
import shutil
from pathlib import Path


DEFAULT_SOURCE = Path("styles/neu-thesis-gbt7714-2015-numeric.csl")
DEFAULT_DEST_DIR = Path.home() / "Zotero" / "styles"


def main() -> None:
    parser = argparse.ArgumentParser(description="Install the project-local NEU thesis CSL into Zotero styles.")
    parser.add_argument("--source", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--dest-dir", type=Path, default=DEFAULT_DEST_DIR)
    args = parser.parse_args()

    if not args.source.exists():
        raise SystemExit(f"Missing CSL source: {args.source}")
    args.dest_dir.mkdir(parents=True, exist_ok=True)
    target = args.dest_dir / args.source.name
    shutil.copyfile(args.source, target)
    print(f"installed_csl={target}")


if __name__ == "__main__":
    main()
