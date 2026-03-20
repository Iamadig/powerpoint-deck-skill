#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from export_pptx_preview import default_work_root, export_preview_images
from pptx_common import slide_page_position


def main() -> None:
    parser = argparse.ArgumentParser(description="Preview one slide from a PPTX into a repo-local work directory.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--slide", type=int, required=True, help="PPT slide number, not page index.")
    parser.add_argument("--output-dir")
    args = parser.parse_args()

    pptx = Path(args.pptx)
    output_dir = Path(args.output_dir) if args.output_dir else default_work_root() / "previews" / pptx.stem
    page = slide_page_position(pptx, args.slide)
    _, images = export_preview_images(pptx, output_dir, slides=[page])
    for image in images:
        print(image)


if __name__ == "__main__":
    main()
