#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import shutil
import tempfile
from pathlib import Path

from classify_slide_archetypes import classify
from clone_slide_archetype import clone_slide
from export_pptx_preview import (
    current_draft_path,
    current_preview_dir,
    export_preview_images,
    named_export_path,
    variant_board_path,
    variant_manifest_path,
    variant_work_dir,
)
from pptx_common import slide_page_position
from safe_slide_composer import compose_deck
from safe_slide_planner import load_content_spec, plan_slide
from synthesize_four_card_grid import synthesize as synthesize_four_card
from variant_board import build_board


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
    from pptx_common import load_slide_summaries

    slides = load_slide_summaries(pptx, default_font_pt=18.0)
    closing_like = [slide.slide_number for slide in slides if classify(slide).archetype == "cover-or-closing"]
    return max(closing_like) if closing_like else None


def candidate_source_slides(pptx: Path, content_json: Path, max_variants: int) -> list[tuple[int, str]]:
    from pptx_common import load_slide_summaries

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
    seen_archetypes: set[str] = set()
    seen_slides: set[int] = set()
    for slide_number, archetype, _, verdict in ranked:
        if verdict != "safe" or slide_number in seen_slides:
            continue
        picks.append((slide_number, archetype))
        seen_slides.add(slide_number)
        seen_archetypes.add(archetype)
        if len(picks) >= max_variants:
            break
    return picks


def build_variant_manifest(
    pptx: Path,
    content_json: Path,
    *,
    max_variants: int = 3,
) -> dict[str, object]:
    output_dir = variant_work_dir(pptx)
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
    return {
        "source_pptx": str(pptx),
        "recommended_variant_id": variants[0]["id"],
        "variants": variants,
    }


def select_variant(manifest: dict[str, object], variant: str | int | None) -> dict[str, object]:
    variants = manifest["variants"]
    if not variants:
        raise SystemExit("no variants available")
    if variant is None:
        choice = manifest.get("recommended_variant_id")
    else:
        choice = str(variant)
        if choice.isdigit():
            choice = f"variant-{choice}"
    selected = next((item for item in variants if item["id"] == choice), None)
    if selected is None:
        available = ", ".join(str(item["id"]) for item in variants)
        raise SystemExit(f"unknown variant '{variant}'; available: {available}")
    return selected


def make_slide(
    pptx: Path,
    content_json: Path,
    *,
    output: Path | None = None,
    preview_dir: Path | None = None,
    save_as: str | None = None,
    source_slide: int | None = None,
    insert_before: int | None = None,
) -> tuple[Path, Path, Path | None]:
    output = output or current_draft_path(pptx)
    preview_dir = preview_dir or current_preview_dir(pptx)
    output.parent.mkdir(parents=True, exist_ok=True)
    preview_dir.mkdir(parents=True, exist_ok=True)
    base_pptx = output if output.exists() else pptx

    payload = json.loads(content_json.read_text())
    spec = load_content_spec(content_json)
    chosen_source = source_slide or best_safe_source_slide(base_pptx, content_json)
    insert_before = insert_before if insert_before is not None else default_insert_before(base_pptx)

    if chosen_source is not None:
        with tempfile.TemporaryDirectory(prefix="ppt-make-slide-") as tmp_dir:
            cloned = Path(tmp_dir) / "cloned.pptx"
            new_slide_number, _ = clone_slide(
                base_pptx,
                source_slide=chosen_source,
                output=cloned,
                insert_before=insert_before,
            )
            compose_deck(cloned, new_slide_number, spec, output, allow_partial=False)
    elif len(payload.get("cards", []) or []) == 4:
        new_slide_number, _ = synthesize_four_card(
            base_pptx,
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
    final_export: Path | None = None
    if save_as:
        final_export = named_export_path(pptx, save_as)
        final_export.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(output, final_export)
    return output, images[0], final_export


def main() -> None:
    parser = argparse.ArgumentParser(description="Create one new slide in an existing deck's design system.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--content-json", required=True)
    parser.add_argument("--output")
    parser.add_argument("--preview-dir")
    parser.add_argument("--save-as", help="Optional named export; working draft still updates in place.")
    parser.add_argument("--source-slide", type=int)
    parser.add_argument("--insert-before", type=int)
    parser.add_argument("--auto", action="store_true", help="Use the recommended variant automatically.")
    parser.add_argument("--variant", help="Choose a generated variant by id or number, e.g. variant-2 or 2.")
    parser.add_argument("--max-variants", type=int, default=3)
    args = parser.parse_args()

    pptx = Path(args.pptx)
    content_json = Path(args.content_json)

    if args.source_slide is None and not args.auto and not args.variant:
        manifest = build_variant_manifest(
            pptx,
            content_json,
            max_variants=args.max_variants,
        )
        manifest_path = variant_manifest_path(pptx)
        manifest_path.parent.mkdir(parents=True, exist_ok=True)
        manifest_path.write_text(json.dumps(manifest, indent=2))
        board_path = variant_board_path(pptx)
        board_path.write_text(build_board(manifest))
        print(f"content: {content_json}")
        print(f"variants: {manifest_path}")
        print(f"board: {board_path}")
        print(f"recommended: {manifest['recommended_variant_id']}")
        print("next: rerun with --auto or --variant <id>")
        return

    chosen_source = args.source_slide
    if chosen_source is None:
        manifest_path = variant_manifest_path(pptx)
        if manifest_path.exists():
            manifest = json.loads(manifest_path.read_text())
        else:
            manifest = build_variant_manifest(
                pptx,
                content_json,
                max_variants=args.max_variants,
            )
            manifest_path.parent.mkdir(parents=True, exist_ok=True)
            manifest_path.write_text(json.dumps(manifest, indent=2))
            board_path = variant_board_path(pptx)
            board_path.write_text(build_board(manifest))
        chosen = select_variant(manifest, args.variant)
        chosen_source = int(chosen["source_slide"])

    output, preview, final_export = make_slide(
        pptx,
        content_json,
        output=Path(args.output) if args.output else None,
        preview_dir=Path(args.preview_dir) if args.preview_dir else None,
        save_as=args.save_as,
        source_slide=chosen_source,
        insert_before=args.insert_before,
    )
    print(f"content: {Path(args.content_json)}")
    print(f"draft: {output}")
    print(f"preview: {preview}")
    if final_export is not None:
        print(f"export: {final_export}")


if __name__ == "__main__":
    main()
