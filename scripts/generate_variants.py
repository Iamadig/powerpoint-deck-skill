#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from classify_slide_archetypes import classify
from export_pptx_preview import current_preview_dir, default_work_root, deck_key
from make_slide import make_slide
from pptx_common import load_slide_summaries
from safe_slide_planner import load_content_spec, plan_slide


def candidate_source_slides(pptx: Path, content_json: Path, max_variants: int) -> list[tuple[int, str]]:
    spec = load_content_spec(content_json)
    slides = load_slide_summaries(pptx, default_font_pt=18.0)
    ranked = sorted(
        (
            (
                slide.slide_number,
                classify(slide).archetype,
                plan_slide(slide, spec).score,
                plan_slide(slide, spec).verdict,
            )
            for slide in slides
        ),
        key=lambda item: (item[3] == "safe", item[2]),
        reverse=True,
    )
    picks: list[tuple[int, str]] = []
    seen: set[int] = set()
    for slide_number, archetype, _, verdict in ranked:
        if verdict != "safe" or slide_number in seen:
            continue
        picks.append((slide_number, archetype))
        seen.add(slide_number)
        if len(picks) >= max_variants:
            break
    return picks


def generate_variants(
    pptx: Path,
    content_json: Path,
    *,
    max_variants: int = 3,
    output_dir: Path | None = None,
) -> dict[str, object]:
    work_root = default_work_root()
    output_dir = output_dir or work_root / "variants" / deck_key(pptx)
    output_dir.mkdir(parents=True, exist_ok=True)

    variants: list[dict[str, object]] = []
    candidates = candidate_source_slides(pptx, content_json, max_variants=max_variants)
    if not candidates:
        raise SystemExit("no safe variants available; shorten content or change slide shape")

    for idx, (source_slide, archetype) in enumerate(candidates, start=1):
        label = f"variant-{idx}"
        deck_out = output_dir / f"{label}.pptx"
        preview_dir = output_dir / f"{label}-preview"
        _, preview, _ = make_slide(
            pptx,
            content_json,
            output=deck_out,
            preview_dir=preview_dir,
            source_slide=source_slide,
        )
        variants.append(
            {
                "id": label,
                "label": f"Variant {idx}",
                "deck": str(deck_out),
                "preview": str(preview),
                "source_slide": source_slide,
                "archetype": archetype,
                "notes": [f"derived from slide {source_slide}", f"archetype={archetype}"],
            }
        )
    return {"source_pptx": str(pptx), "variants": variants}


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a few safe slide variants from one content spec.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--content-json", required=True)
    parser.add_argument("--max-variants", type=int, default=3)
    parser.add_argument("--output")
    parser.add_argument("--output-dir")
    args = parser.parse_args()

    payload = generate_variants(
        Path(args.pptx),
        Path(args.content_json),
        max_variants=args.max_variants,
        output_dir=Path(args.output_dir) if args.output_dir else None,
    )
    if args.output:
        out = Path(args.output)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(payload, indent=2))
        print(out)
        return
    print(json.dumps(payload, indent=2))


if __name__ == "__main__":
    main()
