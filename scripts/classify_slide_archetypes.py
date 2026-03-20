#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from dataclasses import asdict, dataclass
from pathlib import Path

from pptx_common import SlideSummary, ShapeSummary, load_slide_summaries


@dataclass
class ArchetypeSummary:
    slide_number: int
    title: str | None
    archetype: str
    layout_signature: str
    region_count: int
    notes: list[str]


def cluster_positions(values: list[int], threshold: int) -> list[int]:
    if not values:
        return []
    values = sorted(values)
    clusters = [[values[0]]]
    for value in values[1:]:
        if abs(value - clusters[-1][-1]) <= threshold:
            clusters[-1].append(value)
        else:
            clusters.append([value])
    return [sum(cluster) // len(cluster) for cluster in clusters]


def title_shape(slide: SlideSummary) -> ShapeSummary | None:
    if not slide.text_shapes:
        return None
    return max(
        slide.text_shapes,
        key=lambda shape: ((shape.font_pt or 0), -shape.y, len(shape.text)),
    )


def role_guess(shape: ShapeSummary, title: ShapeSummary | None) -> str:
    if title and shape.shape_id == title.shape_id:
        return "title"
    font = shape.font_pt or 0
    if shape.y < 700000 and font <= 12:
        return "section-label"
    if font >= 18:
        return "subhead"
    if font >= 13:
        return "card-title"
    if shape.y > 5600000:
        return "footer"
    return "body"


def classify(slide: SlideSummary) -> ArchetypeSummary:
    title = title_shape(slide)
    content_shapes = [shape for shape in slide.text_shapes if role_guess(shape, title) not in {"section-label", "title"}]
    row_clusters = cluster_positions([shape.y for shape in content_shapes], threshold=550000)
    col_clusters = cluster_positions([shape.x for shape in content_shapes], threshold=2200000)

    wide_shapes = [shape for shape in content_shapes if shape.width_pt >= 300]
    mid_shapes = [shape for shape in content_shapes if 110 <= shape.width_pt < 300]
    small_shapes = [shape for shape in content_shapes if shape.width_pt < 110]

    if len(slide.text_shapes) <= 8:
        archetype = "cover-or-closing"
    elif len(col_clusters) >= 3 and any(role_guess(shape, title) == "footer" for shape in content_shapes):
        archetype = "multi-column-grid-with-footer"
    elif len(col_clusters) == 2 and len(row_clusters) >= 3 and wide_shapes:
        archetype = "two-column-grid-with-banner"
    elif len(col_clusters) == 2 and len(row_clusters) == 2:
        archetype = "two-column-comparison"
    elif len(col_clusters) == 1 and len(wide_shapes) >= 4:
        archetype = "single-column-narrative"
    elif any("/" in shape.text for shape in content_shapes) and any((shape.font_pt or 0) >= 30 for shape in content_shapes):
        archetype = "metric-proof"
    else:
        archetype = "custom"

    notes: list[str] = []
    if len(col_clusters) >= 3:
        notes.append(f"{len(col_clusters)} x-clusters detected")
    if len(row_clusters) >= 3:
        notes.append(f"{len(row_clusters)} y-bands detected")
    if wide_shapes:
        notes.append(f"{len(wide_shapes)} wide text regions")
    if mid_shapes:
        notes.append(f"{len(mid_shapes)} mid-width text regions")
    if small_shapes:
        notes.append(f"{len(small_shapes)} compact label regions")

    return ArchetypeSummary(
        slide_number=slide.slide_number,
        title=slide.title,
        archetype=archetype,
        layout_signature=(
            f"rows={len(row_clusters)} cols={len(col_clusters)} "
            f"wide={len(wide_shapes)} mid={len(mid_shapes)} small={len(small_shapes)}"
        ),
        region_count=len(content_shapes),
        notes=notes,
    )


def render_text(item: ArchetypeSummary) -> str:
    notes = "; ".join(item.notes) if item.notes else "no extra notes"
    return (
        f"slide {item.slide_number}: {item.title or '(untitled)'}\n"
        f"  archetype: {item.archetype}\n"
        f"  signature: {item.layout_signature}\n"
        f"  notes: {notes}"
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Classify PPTX slide archetypes.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--slide", type=int)
    parser.add_argument("--format", choices=["text", "json"], default="text")
    args = parser.parse_args()

    results = [classify(slide) for slide in load_slide_summaries(Path(args.pptx), slide_number=args.slide)]

    if args.format == "json":
        print(json.dumps([asdict(item) for item in results], indent=2))
        return

    for item in results:
        print(render_text(item))


if __name__ == "__main__":
    main()
