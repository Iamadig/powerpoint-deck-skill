#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from export_pptx_preview import current_draft_path, current_preview_dir, default_work_root, deck_key


def main() -> None:
    parser = argparse.ArgumentParser(description="Clean PPTX skill working artifacts.")
    parser.add_argument("--pptx", help="Clean only artifacts for this source deck.")
    parser.add_argument("--all", action="store_true", help="Clean the full .pptx-work directory.")
    args = parser.parse_args()

    work_root = default_work_root()
    if args.all:
        if work_root.exists():
            shutil.rmtree(work_root)
        print(work_root)
        return

    if not args.pptx:
        raise SystemExit("pass --pptx <deck.pptx> or --all")

    source = Path(args.pptx)
    key = deck_key(source)
    targets = [
        work_root / "out" / key,
        work_root / "previews" / key,
        work_root / "cache" / f"{key}.pdf",
    ]
    for target in targets:
        if target.is_dir():
            shutil.rmtree(target)
        elif target.exists():
            target.unlink()
    for target in targets:
        print(target)


if __name__ == "__main__":
    main()
