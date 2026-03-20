#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from export_pptx_preview import current_draft_path, named_export_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Copy the current working draft to a named final PPTX.")
    parser.add_argument("--pptx", required=True, help="Original source deck used to create the draft.")
    parser.add_argument("--name", required=True, help="Final export name, with or without .pptx")
    parser.add_argument("--output")
    args = parser.parse_args()

    source = Path(args.pptx)
    draft = current_draft_path(source)
    if not draft.exists():
        raise SystemExit(f"no working draft found: {draft}")

    output = Path(args.output) if args.output else named_export_path(source, args.name)
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(draft, output)
    print(output)


if __name__ == "__main__":
    main()

