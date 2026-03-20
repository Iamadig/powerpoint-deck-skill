#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from collections import defaultdict
from pathlib import Path

from inspect_pptx import summarize_deck


def extract_design_system(pptx: Path) -> dict[str, object]:
    summary = summarize_deck(pptx)
    deck = summary["deck"]
    grouped: dict[str, list[int]] = defaultdict(list)
    for slide in summary["slides"]:
        grouped[slide["archetype"]].append(slide["slide_number"])
    return {
        "source_pptx": str(pptx),
        "canvas": deck["slide_size"],
        "tokens": {
            "colors": deck["colors"],
            "font_sizes_pt": deck["font_sizes_pt"],
            "x_bands": deck["x_bands"],
            "y_bands": deck["y_bands"],
        },
        "archetypes": [
            {"name": name, "count": len(slides), "slides": slides}
            for name, slides in sorted(grouped.items())
        ],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract a clean design-system summary from a PPTX.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--output")
    args = parser.parse_args()

    payload = extract_design_system(Path(args.pptx))
    if args.output:
        out = Path(args.output)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(payload, indent=2))
        print(out)
        return
    print(json.dumps(payload, indent=2))


if __name__ == "__main__":
    main()
