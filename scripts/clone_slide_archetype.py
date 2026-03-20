#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

ET.register_namespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
ET.register_namespace("p", P_NS)
ET.register_namespace("r", R_NS)


def parse_slide_number(path: str) -> int:
    match = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", path)
    if not match:
        raise ValueError(f"not a slide path: {path}")
    return int(match.group(1))


def parse_rel_target_number(target: str) -> int:
    match = re.fullmatch(r"slides/slide(\d+)\.xml", target)
    if not match:
        raise ValueError(f"not a slide target: {target}")
    return int(match.group(1))


def sorted_slide_numbers(zf: ZipFile) -> list[int]:
    numbers: list[int] = []
    for name in zf.namelist():
        if re.fullmatch(r"ppt/slides/slide\d+\.xml", name):
            numbers.append(parse_slide_number(name))
    return sorted(numbers)


def next_slide_number(zf: ZipFile) -> int:
    numbers = sorted_slide_numbers(zf)
    return (max(numbers) + 1) if numbers else 1


def next_rel_id(rels_root: ET.Element) -> str:
    seen: list[int] = []
    for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        rel_id = rel.get("Id", "")
        if rel_id.startswith("rId") and rel_id[3:].isdigit():
            seen.append(int(rel_id[3:]))
    return f"rId{(max(seen) + 1) if seen else 1}"


def next_slide_id(presentation_root: ET.Element) -> int:
    seen: list[int] = []
    for node in presentation_root.findall(f".//{{{P_NS}}}sldId"):
        slide_id = node.get("id")
        if slide_id and slide_id.isdigit():
            seen.append(int(slide_id))
    return (max(seen) + 1) if seen else 256


def read_xml(zf: ZipFile, name: str) -> ET.Element:
    return ET.fromstring(zf.read(name))


def clone_slide_rels(xml_bytes: bytes) -> bytes:
    root = ET.fromstring(xml_bytes)
    for rel in list(root.findall(f"{{{PKG_REL_NS}}}Relationship")):
        rel_type = rel.get("Type", "")
        if rel_type.endswith("/notesSlide"):
            root.remove(rel)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def add_slide_to_presentation(
    presentation_xml: bytes,
    new_rel_id: str,
    new_slide_id: int,
    insert_index: int | None,
) -> bytes:
    root = ET.fromstring(presentation_xml)
    sld_id_list = root.find(f"{{{P_NS}}}sldIdLst")
    if sld_id_list is None:
        raise SystemExit("presentation missing p:sldIdLst")
    new_node = ET.Element(f"{{{P_NS}}}sldId")
    new_node.set("id", str(new_slide_id))
    new_node.set(f"{{{R_NS}}}id", new_rel_id)

    current = list(sld_id_list)
    if insert_index is None or insert_index >= len(current):
        sld_id_list.append(new_node)
    else:
        sld_id_list.insert(max(insert_index, 0), new_node)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def add_relationship(
    rels_xml: bytes,
    rel_id: str,
    target: str,
) -> bytes:
    root = ET.fromstring(rels_xml)
    node = ET.Element(f"{{{PKG_REL_NS}}}Relationship")
    node.set("Id", rel_id)
    node.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
    node.set("Target", target)
    root.append(node)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def add_content_type_override(content_types_xml: bytes, slide_number: int) -> bytes:
    root = ET.fromstring(content_types_xml)
    part_name = f"/ppt/slides/slide{slide_number}.xml"
    for override in root.findall(f"{{{CT_NS}}}Override"):
        if override.get("PartName") == part_name:
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)
    node = ET.Element(f"{{{CT_NS}}}Override")
    node.set("PartName", part_name)
    node.set("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
    root.append(node)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def slide_order_from_presentation(presentation_xml: bytes, rels_xml: bytes) -> list[int]:
    pres = ET.fromstring(presentation_xml)
    rels = ET.fromstring(rels_xml)
    target_by_rel_id = {
        rel.get("Id"): rel.get("Target")
        for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    order: list[int] = []
    sld_id_list = pres.find(f"{{{P_NS}}}sldIdLst")
    if sld_id_list is None:
        return order
    for node in sld_id_list.findall(f"{{{P_NS}}}sldId"):
        rel_id = node.get(f"{{{R_NS}}}id")
        target = target_by_rel_id.get(rel_id or "")
        if target and re.fullmatch(r"slides/slide\d+\.xml", target):
            order.append(parse_rel_target_number(target))
    return order


def clone_slide(
    pptx: Path,
    source_slide: int,
    output: Path,
    insert_before: int | None = None,
) -> tuple[int, int]:
    with ZipFile(pptx) as src:
        new_slide_number = next_slide_number(src)
        source_part = f"ppt/slides/slide{source_slide}.xml"
        source_rels_part = f"ppt/slides/_rels/slide{source_slide}.xml.rels"

        if source_part not in src.namelist():
            raise SystemExit(f"source slide {source_slide} not found")
        source_slide_xml = src.read(source_part)
        source_rels_xml = src.read(source_rels_part) if source_rels_part in src.namelist() else None

        presentation_xml = src.read("ppt/presentation.xml")
        presentation_rels_xml = src.read("ppt/_rels/presentation.xml.rels")
        content_types_xml = src.read("[Content_Types].xml")
        slide_order = slide_order_from_presentation(presentation_xml, presentation_rels_xml)

        insert_index = None
        if insert_before is not None:
            if insert_before not in slide_order:
                raise SystemExit(f"insert-before slide {insert_before} not found in presentation order")
            insert_index = slide_order.index(insert_before)

        pres_rels_root = ET.fromstring(presentation_rels_xml)
        new_rel_id = next_rel_id(pres_rels_root)
        new_slide_id = next_slide_id(ET.fromstring(presentation_xml))

        new_presentation_xml = add_slide_to_presentation(presentation_xml, new_rel_id, new_slide_id, insert_index)
        new_presentation_rels_xml = add_relationship(
            presentation_rels_xml,
            new_rel_id,
            f"slides/slide{new_slide_number}.xml",
        )
        new_content_types_xml = add_content_type_override(content_types_xml, new_slide_number)

        output.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(output, "w", ZIP_DEFLATED) as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "ppt/presentation.xml":
                    data = new_presentation_xml
                elif item.filename == "ppt/_rels/presentation.xml.rels":
                    data = new_presentation_rels_xml
                elif item.filename == "[Content_Types].xml":
                    data = new_content_types_xml
                dst.writestr(item, data)

            dst.writestr(f"ppt/slides/slide{new_slide_number}.xml", source_slide_xml)
            if source_rels_xml is not None:
                dst.writestr(
                    f"ppt/slides/_rels/slide{new_slide_number}.xml.rels",
                    clone_slide_rels(source_rels_xml),
                )

    return new_slide_number, new_slide_id


def main() -> None:
    parser = argparse.ArgumentParser(description="Clone an existing PPTX slide archetype into a new slide.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--source-slide", type=int, required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--insert-before", type=int, help="Insert clone before this slide number in presentation order.")
    args = parser.parse_args()

    new_slide_number, new_slide_id = clone_slide(
        Path(args.pptx),
        source_slide=args.source_slide,
        output=Path(args.output),
        insert_before=args.insert_before,
    )
    print(
        f"cloned slide {args.source_slide} -> slide {new_slide_number} "
        f"(presentation id {new_slide_id})"
    )


if __name__ == "__main__":
    main()
