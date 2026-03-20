#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import math
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile

from pptx_common import extract_text_shapes, iter_slide_parts


AVG_CHAR_WIDTH = 0.52
LINE_HEIGHT = 1.25


@dataclass
class ShapeInfo:
    shape_id: str
    name: str
    width_pt: float
    height_pt: float
    font_pt: float
    current_text: str = ""


def shape_info_for_slide(zf: ZipFile, slide_number: int) -> list[ShapeInfo]:
    part = None
    for num, name in iter_slide_parts(zf):
        if num == slide_number:
            part = name
            break
    if part is None:
        raise SystemExit(f"slide {slide_number} not found")

    shapes, _ = extract_text_shapes(zf.read(part), default_font_pt=18.0)
    return [
        ShapeInfo(
            shape_id=shape.shape_id,
            name=shape.name,
            width_pt=shape.width_pt,
            height_pt=shape.height_pt,
            font_pt=shape.font_pt or 18.0,
            current_text=shape.text,
        )
        for shape in shapes
    ]


def estimate_fit(shape: ShapeInfo, text: str) -> dict[str, object]:
    effective_width = max(shape.width_pt - (shape.font_pt * 0.8), 1)
    chars_per_line = max(int(effective_width / (shape.font_pt * AVG_CHAR_WIDTH)), 1)
    paragraphs = [segment for segment in text.split("\n") if segment.strip()] or [text]
    estimated_lines = 0
    for paragraph in paragraphs:
        estimated_lines += max(1, math.ceil(len(paragraph) / chars_per_line))
    allowed_lines = max(int(shape.height_pt / (shape.font_pt * LINE_HEIGHT)), 1)
    raw_capacity = chars_per_line * allowed_lines
    calibrated_capacity = max(raw_capacity, int(len(shape.current_text) * 1.1))
    ratio = max(estimated_lines / allowed_lines, len(text) / max(calibrated_capacity, 1))
    if ratio <= 0.95:
        risk = "low"
    elif ratio <= 1.15:
        risk = "medium"
    else:
        risk = "high"
    return {
        "shape_id": shape.shape_id,
        "shape_name": shape.name,
        "current_text_chars": len(shape.current_text),
        "font_pt": shape.font_pt,
        "width_pt": round(shape.width_pt, 2),
        "height_pt": round(shape.height_pt, 2),
        "chars_per_line_est": chars_per_line,
        "estimated_lines": estimated_lines,
        "allowed_lines_est": allowed_lines,
        "raw_capacity_chars_est": raw_capacity,
        "calibrated_capacity_chars_est": calibrated_capacity,
        "risk": risk,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Estimate whether replacement text will fit a PPTX text box.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--slide", required=True, type=int)
    parser.add_argument("--shape-name")
    parser.add_argument("--shape-id")
    parser.add_argument("--text")
    parser.add_argument("--mapping-json", help="JSON file: [{shape_name|shape_id, text}, ...]")
    args = parser.parse_args()

    with ZipFile(Path(args.pptx)) as zf:
        shapes = shape_info_for_slide(zf, args.slide)

    shape_by_name = {shape.name: shape for shape in shapes}
    shape_by_id = {shape.shape_id: shape for shape in shapes}

    requests: list[tuple[ShapeInfo, str]] = []
    if args.mapping_json:
        payload = json.loads(Path(args.mapping_json).read_text())
        for item in payload:
            shape = None
            if item.get("shape_name"):
                shape = shape_by_name.get(item["shape_name"])
            elif item.get("shape_id"):
                shape = shape_by_id.get(str(item["shape_id"]))
            if shape is None:
                raise SystemExit(f"shape not found for mapping item: {item}")
            requests.append((shape, item["text"]))
    else:
        if not args.text:
            raise SystemExit("--text is required unless --mapping-json is used")
        shape = None
        if args.shape_name:
            shape = shape_by_name.get(args.shape_name)
        elif args.shape_id:
            shape = shape_by_id.get(str(args.shape_id))
        if shape is None:
            raise SystemExit("provide --shape-name or --shape-id for single-shape mode")
        requests.append((shape, args.text))

    print(json.dumps([estimate_fit(shape, text) for shape, text in requests], indent=2))


if __name__ == "__main__":
    main()
