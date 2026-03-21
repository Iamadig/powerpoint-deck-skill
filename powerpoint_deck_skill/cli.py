from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = ROOT / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from export_pptx_preview import (  # type: ignore  # noqa: E402
    current_draft_path,
    current_preview_dir,
    default_work_root,
    deck_key,
    export_preview_images,
    named_export_path,
    variant_board_path,
    variant_manifest_path,
)
from extract_design_system import extract_design_system  # type: ignore  # noqa: E402
from inspect_pptx import render_slide_summary, summarize_deck  # type: ignore  # noqa: E402
from make_slide import (  # type: ignore  # noqa: E402
    build_variant_manifest,
    make_slide as render_slide,
    select_variant,
)
from pptx_common import slide_page_position  # type: ignore  # noqa: E402
from variant_board import build_board  # type: ignore  # noqa: E402


def cmd_inspect(args: argparse.Namespace) -> int:
    summary = summarize_deck(Path(args.pptx), slide_number=args.slide)
    if args.format == "json":
        print(json.dumps(summary, indent=2))
        return 0
    deck = summary["deck"]
    print(f"deck: {deck['path']}")
    print(f"slide size: {deck['slide_size']['width_pt']}pt x {deck['slide_size']['height_pt']}pt")
    print(f"colors: text={', '.join(deck['colors']['text'])} fill={', '.join(deck['colors']['fill'])}")
    print(f"font sizes: {', '.join(str(item) for item in deck['font_sizes_pt'])}")
    for slide_data in summary["slides"]:
        print(render_slide_summary(slide_data))
    return 0


def cmd_design_system(args: argparse.Namespace) -> int:
    payload = extract_design_system(Path(args.pptx))
    if args.output:
        out = Path(args.output)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(payload, indent=2))
        print(out)
        return 0
    print(json.dumps(payload, indent=2))
    return 0


def cmd_variants(args: argparse.Namespace) -> int:
    pptx = Path(args.pptx)
    content_json = Path(args.content_json)
    manifest = build_variant_manifest(
        pptx,
        content_json,
        max_variants=args.max_variants,
    )
    manifest_path = Path(args.output) if args.output else variant_manifest_path(pptx)
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(json.dumps(manifest, indent=2))
    board_path = Path(args.board) if args.board else variant_board_path(pptx)
    board_path.parent.mkdir(parents=True, exist_ok=True)
    board_path.write_text(build_board(manifest))
    print(f"variants: {manifest_path}")
    print(f"board: {board_path}")
    print(f"recommended: {manifest['recommended_variant_id']}")
    return 0


def cmd_make(args: argparse.Namespace) -> int:
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
        board_path.parent.mkdir(parents=True, exist_ok=True)
        board_path.write_text(build_board(manifest))
        print(f"content: {content_json}")
        print(f"variants: {manifest_path}")
        print(f"board: {board_path}")
        print(f"recommended: {manifest['recommended_variant_id']}")
        print("next: rerun with --auto or --variant <id>")
        return 0

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
            board_path.parent.mkdir(parents=True, exist_ok=True)
            board_path.write_text(build_board(manifest))
        chosen = select_variant(manifest, args.variant)
        chosen_source = int(chosen["source_slide"])

    output, preview, final_export = render_slide(
        pptx,
        content_json,
        output=Path(args.output) if args.output else None,
        preview_dir=Path(args.preview_dir) if args.preview_dir else None,
        save_as=args.save_as,
        source_slide=chosen_source,
        insert_before=args.insert_before,
    )
    print(f"content: {content_json}")
    print(f"draft: {output}")
    print(f"preview: {preview}")
    if final_export is not None:
        print(f"export: {final_export}")
    return 0


def cmd_preview(args: argparse.Namespace) -> int:
    source_pptx = Path(args.pptx)
    working_pptx = current_draft_path(source_pptx)
    pptx = source_pptx if args.use_source or not working_pptx.exists() else working_pptx
    output_dir = Path(args.output_dir) if args.output_dir else current_preview_dir(source_pptx)
    page = slide_page_position(pptx, args.slide)
    _, images = export_preview_images(pptx, output_dir, slides=[page])
    for image in images:
        print(image)
    return 0


def cmd_export(args: argparse.Namespace) -> int:
    source = Path(args.pptx)
    draft = current_draft_path(source)
    if not draft.exists():
        raise SystemExit(f"no working draft found: {draft}")
    output = Path(args.output) if args.output else named_export_path(source, args.name)
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(draft, output)
    print(output)
    return 0


def cmd_clean(args: argparse.Namespace) -> int:
    work_root = default_work_root()
    if args.all:
        if work_root.exists():
            shutil.rmtree(work_root)
        print(work_root)
        return 0
    if not args.pptx:
        raise SystemExit("pass --pptx <deck.pptx> or --all")
    source = Path(args.pptx)
    key = deck_key(source)
    targets = [
        work_root / "out" / key,
        work_root / "previews" / key,
        work_root / "variants" / key,
        work_root / "cache" / f"{key}.pdf",
    ]
    for target in targets:
        if target.is_dir():
            shutil.rmtree(target)
        elif target.exists():
            target.unlink()
    for target in targets:
        print(target)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="pptx-skill", description="Inspect decks and generate slides in the same design system.")
    sub = parser.add_subparsers(dest="command", required=True)

    inspect_parser = sub.add_parser("inspect", help="Inspect a deck's content and geometry.")
    inspect_parser.add_argument("pptx")
    inspect_parser.add_argument("--slide", type=int)
    inspect_parser.add_argument("--format", choices=["text", "json"], default="text")
    inspect_parser.set_defaults(func=cmd_inspect)

    ds_parser = sub.add_parser("design-system", help="Extract a clean design-system summary.")
    ds_parser.add_argument("pptx")
    ds_parser.add_argument("--output")
    ds_parser.set_defaults(func=cmd_design_system)

    variants_parser = sub.add_parser("variants", help="Generate variants and an HTML board.")
    variants_parser.add_argument("pptx")
    variants_parser.add_argument("--content-json", required=True)
    variants_parser.add_argument("--max-variants", type=int, default=3)
    variants_parser.add_argument("--output")
    variants_parser.add_argument("--board")
    variants_parser.set_defaults(func=cmd_variants)

    make_parser = sub.add_parser("make", help="Generate options first; apply only on explicit choice.")
    make_parser.add_argument("pptx")
    make_parser.add_argument("--content-json", required=True)
    make_parser.add_argument("--output")
    make_parser.add_argument("--preview-dir")
    make_parser.add_argument("--save-as")
    make_parser.add_argument("--source-slide", type=int)
    make_parser.add_argument("--insert-before", type=int)
    make_parser.add_argument("--auto", action="store_true")
    make_parser.add_argument("--variant")
    make_parser.add_argument("--max-variants", type=int, default=3)
    make_parser.set_defaults(func=cmd_make)

    preview_parser = sub.add_parser("preview", help="Preview one slide from source or current draft.")
    preview_parser.add_argument("pptx")
    preview_parser.add_argument("--slide", type=int, required=True)
    preview_parser.add_argument("--output-dir")
    preview_parser.add_argument("--use-source", action="store_true")
    preview_parser.set_defaults(func=cmd_preview)

    export_parser = sub.add_parser("export", help="Copy current draft to a named final PPTX.")
    export_parser.add_argument("pptx")
    export_parser.add_argument("--name", required=True)
    export_parser.add_argument("--output")
    export_parser.set_defaults(func=cmd_export)

    clean_parser = sub.add_parser("clean", help="Clean working artifacts.")
    clean_parser.add_argument("pptx", nargs="?")
    clean_parser.add_argument("--all", action="store_true")
    clean_parser.set_defaults(func=cmd_clean)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return int(args.func(args))
