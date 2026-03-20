#!/usr/bin/env python3
from __future__ import annotations

import argparse
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from classify_slide_archetypes import classify
from clone_slide_archetype import clone_slide
from export_pptx_preview import export_preview_images
from layout_engine import banner_with_callout, grid_2x2
from pptx_common import NS, load_presentation_size, load_slide_summaries


def find_shape_by_name(root: ET.Element, name: str) -> ET.Element:
    for sp in root.findall(".//p:sp", NS):
        nv = sp.find("p:nvSpPr/p:cNvPr", NS)
        if nv is not None and nv.get("name") == name:
            return sp
    raise SystemExit(f"shape not found: {name}")


def set_geometry(sp: ET.Element, x: int, y: int, cx: int, cy: int) -> None:
    xfrm = sp.find("p:spPr/a:xfrm", NS)
    if xfrm is None:
        return
    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return
    off.set("x", str(x))
    off.set("y", str(y))
    ext.set("cx", str(cx))
    ext.set("cy", str(cy))


def set_text(sp: ET.Element, text: str) -> None:
    nodes = sp.findall(".//a:t", NS)
    if not nodes:
        return
    nodes[0].text = text
    for node in nodes[1:]:
        node.text = ""


def set_font_size(sp: ET.Element, pt: float) -> None:
    size = str(int(pt * 100))
    for node in sp.findall(".//a:rPr", NS):
        node.set("sz", size)
    for node in sp.findall(".//a:endParaRPr", NS):
        node.set("sz", size)
    for node in sp.findall(".//a:defRPr", NS):
        node.set("sz", size)


def rewrite_slide(xml_bytes: bytes, payload: dict[str, object]) -> bytes:
    root = ET.fromstring(xml_bytes)
    hide_banner = bool(payload.get("hide_banner"))

    text_map = {
        "Text 1": payload["section_label"],
        "Text 2": payload["title"],
        "Text 42": "" if hide_banner else payload["banner_title"],
        "Text 43": "" if hide_banner else payload["banner_body"],
        "Text 46": "" if hide_banner else payload["callout"],
    }
    for name, text in text_map.items():
        set_text(find_shape_by_name(root, name), str(text))

    slide_w = int(payload.get("_slide_cx", 12192000))
    slide_h = int(payload.get("_slide_cy", 6858000))
    banner = banner_with_callout(
        slide_w,
        left=1300162,
        right=285750,
        top=1590000,
        height=850000,
        callout_w=2700000,
        gap=250000,
    )
    if hide_banner:
        offscreen_x = slide_w + 500000
        for shape_name in ["Shape 39", "Shape 40", "Shape 41", "Text 42", "Text 43", "Shape 44", "Shape 45", "Text 46"]:
            set_geometry(find_shape_by_name(root, shape_name), offscreen_x, 0, 1000, 1000)
    else:
        set_geometry(find_shape_by_name(root, "Text 42"), banner["body"].x, 1590000, banner["body"].cx, 266700)
        set_geometry(find_shape_by_name(root, "Text 43"), banner["body"].x, 1960000, banner["body"].cx, 533400)
        set_geometry(find_shape_by_name(root, "Text 46"), banner["callout"].x, banner["callout"].y, banner["callout"].cx, banner["callout"].cy)
        set_font_size(find_shape_by_name(root, "Text 42"), 15.0)
        set_font_size(find_shape_by_name(root, "Text 43"), 11.5)
        set_font_size(find_shape_by_name(root, "Text 46"), 10.5)

    cards = payload["cards"]
    slots = [
        ("Shape 3", "Shape 4", "Shape 5", "Text 6", "Text 7", "Text 9", "Text 11", "Shape 8", "Shape 10"),
        ("Shape 12", "Shape 13", "Shape 14", "Text 15", "Text 16", "Text 18", "Text 20", "Shape 17", "Shape 19"),
        ("Shape 21", "Shape 22", "Shape 23", "Text 24", "Text 25", "Text 27", "Text 29", "Shape 26", "Shape 28"),
        ("Shape 30", "Shape 31", "Shape 32", "Text 33", "Text 34", "Text 36", "Text 38", "Shape 35", "Shape 37"),
    ]

    card_bases = [(item.x, item.y) for item in grid_2x2(
        slide_w,
        slide_h,
        left=285750,
        right=285750,
        top=2200000 if hide_banner else 3236002,
        bottom=550000,
        gap_x=260000,
        gap_y=250000,
    )]

    for (
        container_name,
        icon_bg_name,
        icon_glyph_name,
        title_name,
        quote_name,
        pain_name,
        value_name,
        positive_icon,
        negative_icon,
    ), card, (base_x, base_y) in zip(slots, cards, card_bases):
        set_text(find_shape_by_name(root, title_name), card["title"])
        set_text(find_shape_by_name(root, quote_name), card["quote"])
        set_text(find_shape_by_name(root, pain_name), card["pain"])
        set_text(find_shape_by_name(root, value_name), card["value"])

        set_geometry(find_shape_by_name(root, container_name), base_x, base_y, 5610225, 1533525)
        set_geometry(find_shape_by_name(root, icon_bg_name), base_x + 233362, base_y + 233362, 457200, 457200)
        set_geometry(find_shape_by_name(root, icon_glyph_name), base_x + 378618, base_y + 366712, 190500, 190500)
        # Wider, stacked text layout inside each card.
        set_geometry(find_shape_by_name(root, title_name), base_x + 842962, base_y + 233362, 3600000, 266700)
        set_geometry(find_shape_by_name(root, quote_name), base_x + 842962, base_y + 500062, 3900000, 190500)
        set_geometry(find_shape_by_name(root, pain_name), base_x + 476250, base_y + 842962, 4500000, 190500)
        set_geometry(find_shape_by_name(root, value_name), base_x + 476250, base_y + 1109620, 4500000, 190500)
        # Existing line icons are positive-on-top / negative-below. Swap so pain aligns to negative.
        set_geometry(find_shape_by_name(root, positive_icon), base_x + 269081, base_y + 1105362, 116681, 133350)
        set_geometry(find_shape_by_name(root, negative_icon), base_x + 277416, base_y + 837612, 100013, 133350)

        set_font_size(find_shape_by_name(root, title_name), 14.0)
        set_font_size(find_shape_by_name(root, quote_name), 10.5)
        set_font_size(find_shape_by_name(root, pain_name), 10.0)
        set_font_size(find_shape_by_name(root, value_name), 10.0)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def synthesize(
    pptx: Path,
    output: Path,
    payload: dict[str, object],
    *,
    insert_before: int | None = None,
) -> tuple[int, Path]:
    with tempfile.TemporaryDirectory(prefix="pptx-four-card-") as tmp_dir:
        tmp_clone = Path(tmp_dir) / "cloned.pptx"
        new_slide_number, _ = clone_slide(
            pptx,
            source_slide=5,
            output=tmp_clone,
            insert_before=insert_before,
        )
        target_part = f"ppt/slides/slide{new_slide_number}.xml"
        size = load_presentation_size(pptx)
        payload = dict(payload)
        payload["_slide_cx"] = size.cx
        payload["_slide_cy"] = size.cy
        with ZipFile(tmp_clone) as src:
            items = [(item, src.read(item.filename)) for item in src.infolist()]
        output.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(output, "w", ZIP_DEFLATED) as dst:
            for item, data in items:
                if item.filename == target_part:
                    data = rewrite_slide(data, payload)
                dst.writestr(item, data)
    return new_slide_number, output


def default_insert_before(pptx: Path) -> int | None:
    slides = load_slide_summaries(pptx, default_font_pt=18.0)
    closing_like = [slide.slide_number for slide in slides if classify(slide).archetype == "cover-or-closing"]
    if closing_like:
        return max(closing_like)
    if slides:
        return slides[-1].slide_number
    return None


def slide_page_position(pptx: Path, slide_number: int) -> int:
    p_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pkg_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    with ZipFile(pptx) as zf:
        pres = ET.fromstring(zf.read("ppt/presentation.xml"))
        rels = ET.fromstring(zf.read("ppt/_rels/presentation.xml.rels"))
    target_by_rel_id = {
        rel.get("Id"): rel.get("Target")
        for rel in rels.findall(f"{{{pkg_ns}}}Relationship")
    }
    ordered_targets: list[str] = []
    for node in pres.find(f"{{{p_ns}}}sldIdLst").findall(f"{{{p_ns}}}sldId"):
        rel_id = node.get(f"{{{r_ns}}}id")
        target = target_by_rel_id.get(rel_id or "")
        if target:
            ordered_targets.append(target)
    target = f"slides/slide{slide_number}.xml"
    return ordered_targets.index(target) + 1


def main() -> None:
    parser = argparse.ArgumentParser(description="Synthesize a new 4-card audience/use-case slide in the deck's house style.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--content-json", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--insert-before", type=int)
    parser.add_argument("--preview-dir")
    args = parser.parse_args()

    import json

    payload = json.loads(Path(args.content_json).read_text())
    output = Path(args.output)
    slide_number, _ = synthesize(
        Path(args.pptx),
        output,
        payload,
        insert_before=args.insert_before if args.insert_before is not None else default_insert_before(Path(args.pptx)),
    )
    print(f"synthesized 4-card slide {slide_number} into {output}")
    if args.preview_dir:
        _, images = export_preview_images(output, Path(args.preview_dir), slides=[slide_page_position(output, slide_number)])
        for image in images:
            print(image)


if __name__ == "__main__":
    main()
