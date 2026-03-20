#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from export_pptx_preview import current_draft_path, current_preview_dir, export_preview_images
from pptx_common import slide_page_position


def main() -> None:
    parser = argparse.ArgumentParser(description="Preview one slide from a PPTX into a repo-local work directory.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--slide", type=int, required=True, help="PPT slide number, not page index.")
    parser.add_argument("--output-dir")
    parser.add_argument("--use-source", action="store_true", help="Preview the source deck even if a working draft exists.")
    args = parser.parse_args()

    source_pptx = Path(args.pptx)
    working_pptx = current_draft_path(source_pptx)
    pptx = source_pptx if args.use_source or not working_pptx.exists() else working_pptx
    output_dir = Path(args.output_dir) if args.output_dir else current_preview_dir(source_pptx)
    page = slide_page_position(pptx, args.slide)
    _, images = export_preview_images(pptx, output_dir, slides=[page])
    for image in images:
        print(image)


if __name__ == "__main__":
    main()
