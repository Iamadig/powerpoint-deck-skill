#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from dataclasses import asdict
from pathlib import Path

from classify_slide_archetypes import classify
from pptx_common import NS, SlideSummary, load_presentation_size, load_slide_summaries


def cluster(values: list[int], threshold: int) -> list[int]:
    if not values:
        return []
    values = sorted(values)
    groups = [[values[0]]]
    for value in values[1:]:
        if abs(value - groups[-1][-1]) <= threshold:
            groups[-1].append(value)
        else:
            groups.append([value])
    return [sum(group) // len(group) for group in groups]


def extract_colors(pptx: Path) -> dict[str, list[str]]:
    from collections import Counter
    from xml.etree import ElementTree as ET
    from zipfile import ZipFile

    def color_token(node):
        if node is None:
            return None
        srgb = node.find("a:srgbClr", NS)
        if srgb is not None:
            return f"#{srgb.get('val', '').lower()}"
        return None

    text_colors = Counter()
    fill_colors = Counter()
    with ZipFile(pptx) as zf:
        for name in zf.namelist():
            if not name.startswith("ppt/slides/slide") or not name.endswith(".xml"):
                continue
            root = ET.fromstring(zf.read(name))
            for sp in root.findall(".//p:sp", NS):
                for node in sp.findall(".//a:rPr", NS):
                    token = color_token(node.find("a:solidFill", NS))
                    if token:
                        text_colors[token] += 1
                token = color_token(sp.find("p:spPr/a:solidFill", NS))
                if token:
                    fill_colors[token] += 1
    return {
        "text": [item for item, _ in text_colors.most_common(6)],
        "fill": [item for item, _ in fill_colors.most_common(6)],
    }


def summarize_deck(pptx: Path, slide_number: int | None = None) -> dict[str, object]:
    slides = load_slide_summaries(pptx, slide_number=slide_number, default_font_pt=18.0)
    size = load_presentation_size(pptx)
    all_fonts = sorted({round(shape.font_pt or 0, 1) for slide in slides for shape in slide.text_shapes if shape.font_pt})
    xs = [shape.x for slide in slides for shape in slide.text_shapes]
    ys = [shape.y for slide in slides for shape in slide.text_shapes]
    return {
        "deck": {
            "path": str(pptx),
            "slide_size": {
                "cx": size.cx,
                "cy": size.cy,
                "width_pt": round(size.width_pt, 2),
                "height_pt": round(size.height_pt, 2),
            },
            "colors": extract_colors(pptx),
            "font_sizes_pt": all_fonts,
            "x_bands": cluster(xs, 500000)[:8],
            "y_bands": cluster(ys, 500000)[:8],
        },
        "slides": [
            {
                "slide_number": slide.slide_number,
                "title": slide.title,
                "archetype": classify(slide).archetype,
                "shape_count": slide.shape_count,
                "text_shapes": [asdict(shape) for shape in slide.text_shapes],
            }
            for slide in slides
        ],
    }


def render_text(slide: SlideSummary) -> str:
    lines = [
        f"slide {slide.slide_number}: {slide.title or '(untitled)'}",
        f"  text shapes: {len(slide.text_shapes)} / total shapes: {slide.shape_count}",
    ]
    for shape in slide.text_shapes:
        font = f"{shape.font_pt:.1f}pt" if shape.font_pt else "?"
        geom = f"({shape.x},{shape.y},{shape.cx},{shape.cy})"
        lines.append(f"  - [{shape.shape_id}] {shape.name} {font} {geom} :: {shape.text}")
    return "\n".join(lines)


def render_slide_summary(slide_data: dict[str, object]) -> str:
    lines = [
        f"slide {slide_data['slide_number']}: {slide_data['title'] or '(untitled)'}",
        f"  archetype: {slide_data['archetype']}",
        f"  text shapes: {len(slide_data['text_shapes'])} / total shapes: {slide_data['shape_count']}",
    ]
    for shape in slide_data["text_shapes"]:
        font = f"{shape['font_pt']:.1f}pt" if shape["font_pt"] else "?"
        geom = f"({shape['x']},{shape['y']},{shape['cx']},{shape['cy']})"
        lines.append(f"  - [{shape['shape_id']}] {shape['name']} {font} {geom} :: {shape['text']}")
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract a PowerPoint deck's design system and slide content.")
    parser.add_argument("--pptx", required=True, help="Path to .pptx")
    parser.add_argument("--slide", type=int, help="Inspect one slide only")
    parser.add_argument("--format", choices=["text", "json"], default="text")
    args = parser.parse_args()

    summary = summarize_deck(Path(args.pptx), slide_number=args.slide)

    if args.format == "json":
        print(json.dumps(summary, indent=2))
        return

    deck = summary["deck"]
    print(f"deck: {deck['path']}")
    print(f"slide size: {deck['slide_size']['width_pt']}pt x {deck['slide_size']['height_pt']}pt")
    print(f"colors: text={', '.join(deck['colors']['text'])} fill={', '.join(deck['colors']['fill'])}")
    print(f"font sizes: {', '.join(str(item) for item in deck['font_sizes_pt'])}")
    for slide_data in summary["slides"]:
        print(render_slide_summary(slide_data))


if __name__ == "__main__":
    main()
