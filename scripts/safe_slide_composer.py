#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from pptx_common import NS, iter_slide_parts, load_slide_summaries
from safe_slide_planner import (
    ContentSpec,
    ShapeSlot,
    assign_many_with_preferences,
    assign_single_with_preferences,
    flatten_card_bodies,
    flatten_card_meta,
    flatten_card_text,
    load_content_spec,
    plan_slide,
    slide_slots,
)


MANAGED_ROLES = {
    "section-label",
    "title",
    "subtitle",
    "lead",
    "card-title",
    "card-body",
    "card-meta",
    "footer",
}


def slide_part_name(pptx: Path, slide_number: int) -> str:
    with ZipFile(pptx) as zf:
        for num, part in iter_slide_parts(zf):
            if num == slide_number:
                return part
    raise SystemExit(f"slide {slide_number} not found")


def build_assignments(slots: list[ShapeSlot], spec: ContentSpec) -> list[tuple[ShapeSlot, str]]:
    assignments: list[tuple[ShapeSlot, str]] = []
    used: set[str] = set()

    def unused() -> list[ShapeSlot]:
        return [item for item in slots if item.shape_id not in used]

    def assign_in_visual_order(candidates: list[ShapeSlot], texts: list[str], label: str) -> list[tuple[ShapeSlot, str]]:
        ordered_slots = sorted(candidates, key=lambda item: (item.y, item.x))
        if len(ordered_slots) < len(texts):
            raise SystemExit(f"not enough {label} slots")
        pairs: list[tuple[ShapeSlot, str]] = []
        for slot, text in zip(ordered_slots, texts):
            if slot.safe_chars_est < len(text):
                raise SystemExit(
                    f"{label} text too long for visual slot order: len={len(text)} cap={slot.safe_chars_est}"
                )
            pairs.append((slot, text))
        return pairs

    singletons = [
        ([{"section-label"}], spec.section_label),
        ([{"title"}], spec.title),
        ([{"subtitle"}], spec.subtitle),
        ([{"lead"}, {"body"}], spec.lead),
    ]
    for groups, text in singletons:
        for slot in assign_single_with_preferences(unused(), text, groups):
            assignments.append((slot, text))
            used.add(slot.shape_id)

    card_texts = flatten_card_text(spec.cards or [])
    for slot, text in assign_in_visual_order(
        [item for item in unused() if item.role == "card-title"],
        card_texts,
        "card heading",
    ):
        assignments.append((slot, text))
        used.add(slot.shape_id)

    card_bodies = flatten_card_bodies(spec.cards or [])
    for slot, text in assign_in_visual_order(
        [item for item in unused() if item.role == "card-body"],
        card_bodies,
        "card body",
    ):
        assignments.append((slot, text))
        used.add(slot.shape_id)

    card_meta = flatten_card_meta(spec.cards or [])
    for slot, text in assign_in_visual_order(
        [item for item in unused() if item.role == "card-meta"],
        card_meta,
        "card meta",
    ):
        assignments.append((slot, text))
        used.add(slot.shape_id)

    for slot in assign_single_with_preferences(unused(), spec.footer, [{"footer"}, {"body"}]):
        assignments.append((slot, spec.footer))
        used.add(slot.shape_id)

    return sorted(assignments, key=lambda item: (item[0].y, item[0].x))


def set_shape_text(sp: ET.Element, text: str) -> None:
    text_nodes = sp.findall(".//a:t", NS)
    if not text_nodes:
        return
    text_nodes[0].text = text
    for node in text_nodes[1:]:
        node.text = ""


def move_shape_off_canvas(sp: ET.Element) -> None:
    xfrm = sp.find("p:spPr/a:xfrm", NS)
    if xfrm is None:
        return
    off = xfrm.find("a:off", NS)
    if off is None:
        return
    off.set("x", "20000000")
    off.set("y", "20000000")


def apply_assignments_to_slide_xml(
    xml_bytes: bytes,
    assignments: list[tuple[ShapeSlot, str]],
    all_slots: list[ShapeSlot],
    *,
    hide_footer: bool = False,
) -> bytes:
    root = ET.fromstring(xml_bytes)
    by_id = {slot.shape_id: text for slot, text in assignments}
    managed_ids = {slot.shape_id for slot in all_slots if slot.role in MANAGED_ROLES}
    for sp in root.findall(".//p:sp", NS):
        nv = sp.find("p:nvSpPr/p:cNvPr", NS)
        if nv is None:
            continue
        shape_id = nv.get("id", "")
        if shape_id in by_id:
            set_shape_text(sp, by_id[shape_id])
        elif shape_id in managed_ids:
            set_shape_text(sp, "")
        if hide_footer:
            xfrm = sp.find("p:spPr/a:xfrm", NS)
            if xfrm is not None:
                off = xfrm.find("a:off", NS)
                if off is not None and int(off.get("y", "0")) >= 5900000:
                    set_shape_text(sp, "")
                    move_shape_off_canvas(sp)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def render_assignment(slot: ShapeSlot, text: str) -> str:
    return (
        f"  - {slot.shape_name} [{slot.role}] cap≈{slot.safe_chars_est} "
        f"<- {json.dumps(text)}"
    )


def compose_deck(
    pptx: Path,
    slide_number: int,
    spec: ContentSpec,
    output: Path,
    *,
    allow_partial: bool = False,
) -> tuple[object, list[tuple[ShapeSlot, str]]]:
    slide = load_slide_summaries(pptx, slide_number=slide_number, default_font_pt=18.0)[0]
    plan = plan_slide(slide, spec)
    slots = slide_slots(slide)
    if plan.verdict != "safe" and not (allow_partial and plan.verdict == "partial"):
        raise SystemExit(f"refused: slide {slide_number} verdict={plan.verdict}")

    assignments = build_assignments(slots, spec)
    target_part = slide_part_name(pptx, slide_number)
    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(pptx) as src:
        items = [(item, src.read(item.filename)) for item in src.infolist()]
    with ZipFile(output, "w", ZIP_DEFLATED) as dst:
        for item, data in items:
            if item.filename == target_part:
                data = apply_assignments_to_slide_xml(
                    data,
                    assignments,
                    slots,
                    hide_footer=spec.hide_footer,
                )
            dst.writestr(item, data)
    return plan, assignments


def main() -> None:
    parser = argparse.ArgumentParser(description="Safely compose PPTX slide copy only when planner says the fit is safe.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--content-json", required=True)
    parser.add_argument("--slide", type=int, required=True, help="Target existing slide number.")
    parser.add_argument("--output", help="Write edited deck to this path.")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--allow-partial", action="store_true", help="Allow planner verdict=partial.")
    args = parser.parse_args()

    pptx = Path(args.pptx)
    spec = load_content_spec(Path(args.content_json))
    slide = load_slide_summaries(pptx, slide_number=args.slide, default_font_pt=18.0)[0]
    plan = plan_slide(slide, spec)
    slots = slide_slots(slide)

    if plan.verdict != "safe" and not (args.allow_partial and plan.verdict == "partial"):
        print(f"refused: slide {args.slide} verdict={plan.verdict}")
        if plan.warnings:
            for warning in plan.warnings:
                print(f"  warning: {warning}")
        raise SystemExit(2)

    assignments = build_assignments(slots, spec)
    print(f"compose slide {args.slide}: verdict={plan.verdict} archetype={plan.archetype}")
    for slot, text in assignments:
        print(render_assignment(slot, text))

    if args.dry_run:
        return
    if not args.output:
        raise SystemExit("--output required unless --dry-run is set")
    compose_deck(pptx, args.slide, spec, Path(args.output), allow_partial=args.allow_partial)


if __name__ == "__main__":
    main()
