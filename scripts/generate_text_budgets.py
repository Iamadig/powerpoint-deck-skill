#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from estimate_text_fit import ShapeInfo, estimate_fit
from pptx_common import ShapeSummary, load_slide_summaries


def title_id(shapes: list[ShapeSummary]) -> str | None:
    if not shapes:
        return None
    title = max(shapes, key=lambda shape: ((shape.font_pt or 0), -shape.y, len(shape.text)))
    return title.shape_id


def role_guess(shape: ShapeSummary, title_shape_id: str | None) -> str:
    font = shape.font_pt or 0
    if title_shape_id and shape.shape_id == title_shape_id:
        return "title"
    if shape.y < 700000 and font <= 12:
        return "section-label"
    if font >= 18:
        return "subhead"
    if font >= 13:
        return "card-title"
    if shape.y > 5600000:
        return "footer"
    return "body"


def budget_for_shape(shape: ShapeSummary, role: str) -> dict[str, object]:
    fit = estimate_fit(
        ShapeInfo(
            shape_id=shape.shape_id,
            name=shape.name,
            width_pt=shape.width_pt,
            height_pt=shape.height_pt,
            font_pt=shape.font_pt or 18.0,
        ),
        shape.text,
    )
    return {
        "shape_id": shape.shape_id,
        "shape_name": shape.name,
        "role": role,
        "current_text": shape.text,
        "current_chars": len(shape.text),
        "current_words": len(shape.text.split()),
        "font_pt": shape.font_pt,
        "width_pt": round(shape.width_pt, 2),
        "height_pt": round(shape.height_pt, 2),
        "chars_per_line_est": fit["chars_per_line_est"],
        "allowed_lines_est": fit["allowed_lines_est"],
        "raw_capacity_chars_est": fit["raw_capacity_chars_est"],
        "safe_chars_est": fit["calibrated_capacity_chars_est"],
        "occupancy_ratio": round(
            (len(shape.text) / fit["calibrated_capacity_chars_est"]) if fit["calibrated_capacity_chars_est"] else 0,
            2,
        ),
        "current_risk": fit["risk"],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate copy budgets for PPTX text boxes.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--slide", type=int)
    parser.add_argument("--format", choices=["text", "json"], default="text")
    args = parser.parse_args()

    slides = load_slide_summaries(Path(args.pptx), slide_number=args.slide, default_font_pt=18.0)
    payload = []
    for slide in slides:
        slide_title_id = title_id(slide.text_shapes)
        budgets = [budget_for_shape(shape, role_guess(shape, slide_title_id)) for shape in slide.text_shapes]
        payload.append({"slide_number": slide.slide_number, "title": slide.title, "budgets": budgets})

    if args.format == "json":
        print(json.dumps(payload, indent=2))
        return

    for slide in payload:
        print(f"slide {slide['slide_number']}: {slide['title'] or '(untitled)'}")
        for item in slide["budgets"]:
            print(
                f"  - {item['shape_name']} [{item['role']}] "
                f"safe≈{item['safe_chars_est']} chars "
                f"({item['allowed_lines_est']} lines @ {item['chars_per_line_est']} c/line) "
                f"current={item['current_chars']} chars risk={item['current_risk']}"
            )


if __name__ == "__main__":
    main()
