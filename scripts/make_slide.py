#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import tempfile
from pathlib import Path

from clone_slide_archetype import clone_slide
from export_pptx_preview import default_work_root, export_preview_images
from pptx_common import slide_page_position
from safe_slide_composer import compose_deck
from safe_slide_planner import load_content_spec, plan_slide
from synthesize_four_card_grid import synthesize as synthesize_four_card


def best_safe_source_slide(pptx: Path, content_json: Path) -> int | None:
    from pptx_common import load_slide_summaries

    spec = load_content_spec(content_json)
    slides = load_slide_summaries(pptx, default_font_pt=18.0)
    plans = sorted(
        (plan_slide(slide, spec) for slide in slides),
        key=lambda item: (item.verdict == "safe", item.score),
        reverse=True,
    )
    safe = next((item for item in plans if item.verdict == "safe"), None)
    return safe.slide_number if safe else None


def default_insert_before(pptx: Path) -> int | None:
    from classify_slide_archetypes import classify
    from pptx_common import load_slide_summaries

    slides = load_slide_summaries(pptx, default_font_pt=18.0)
    closing_like = [slide.slide_number for slide in slides if classify(slide).archetype == "cover-or-closing"]
    return max(closing_like) if closing_like else None


def make_slide(
    pptx: Path,
    content_json: Path,
    *,
    output: Path | None = None,
    preview_dir: Path | None = None,
    source_slide: int | None = None,
    insert_before: int | None = None,
) -> tuple[Path, Path]:
    work_root = default_work_root()
    output = output or work_root / "out" / f"{pptx.stem}-edited.pptx"
    preview_dir = preview_dir or work_root / "previews" / output.stem
    output.parent.mkdir(parents=True, exist_ok=True)
    preview_dir.mkdir(parents=True, exist_ok=True)

    payload = json.loads(content_json.read_text())
    spec = load_content_spec(content_json)
    chosen_source = source_slide or best_safe_source_slide(pptx, content_json)
    insert_before = insert_before if insert_before is not None else default_insert_before(pptx)

    if chosen_source is not None:
        with tempfile.TemporaryDirectory(prefix="ppt-make-slide-") as tmp_dir:
            cloned = Path(tmp_dir) / "cloned.pptx"
            new_slide_number, _ = clone_slide(
                pptx,
                source_slide=chosen_source,
                output=cloned,
                insert_before=insert_before,
            )
            compose_deck(cloned, new_slide_number, spec, output, allow_partial=False)
    elif len(payload.get("cards", []) or []) == 4:
        new_slide_number, _ = synthesize_four_card(
            pptx,
            output,
            payload,
            insert_before=insert_before,
        )
    else:
        raise SystemExit(
            "no safe existing slide fit; simplify the content or provide a 4-card payload for synthesis"
        )

    page = slide_page_position(output, new_slide_number)
    _, images = export_preview_images(output, preview_dir, slides=[page])
    if not images:
        raise SystemExit("preview export failed")
    return output, images[0]


def main() -> None:
    parser = argparse.ArgumentParser(description="Create one new slide in an existing deck's design system.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--content-json", required=True)
    parser.add_argument("--output")
    parser.add_argument("--preview-dir")
    parser.add_argument("--source-slide", type=int)
    parser.add_argument("--insert-before", type=int)
    args = parser.parse_args()

    output, preview = make_slide(
        Path(args.pptx),
        Path(args.content_json),
        output=Path(args.output) if args.output else None,
        preview_dir=Path(args.preview_dir) if args.preview_dir else None,
        source_slide=args.source_slide,
        insert_before=args.insert_before,
    )
    print(f"content: {Path(args.content_json)}")
    print(f"deck: {output}")
    print(f"preview: {preview}")


if __name__ == "__main__":
    main()
