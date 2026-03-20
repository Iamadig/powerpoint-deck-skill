#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from zipfile import ZipFile


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

EMU_PER_POINT = 12700


@dataclass
class ShapeSummary:
    shape_id: str
    name: str
    x: int
    y: int
    cx: int
    cy: int
    text: str
    font_pt: float | None

    @property
    def width_pt(self) -> float:
        return self.cx / EMU_PER_POINT

    @property
    def height_pt(self) -> float:
        return self.cy / EMU_PER_POINT


@dataclass
class SlideSummary:
    slide_number: int
    title: str | None
    shape_count: int
    text_shapes: list[ShapeSummary]


@dataclass
class PresentationSize:
    cx: int
    cy: int

    @property
    def width_pt(self) -> float:
        return self.cx / EMU_PER_POINT

    @property
    def height_pt(self) -> float:
        return self.cy / EMU_PER_POINT


def iter_slide_parts(zf: ZipFile) -> Iterable[tuple[int, str]]:
    parts: list[tuple[int, str]] = []
    for name in zf.namelist():
        match = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", name)
        if match:
            parts.append((int(match.group(1)), name))
    yield from sorted(parts)


def text_of_shape(sp: ET.Element) -> str:
    parts = [node.text for node in sp.findall(".//a:t", NS) if node.text]
    return " ".join(part.strip() for part in parts if part.strip())


def first_font_pt(sp: ET.Element, *, default: float | None = None) -> float | None:
    for node in sp.findall(".//a:rPr", NS):
        size = node.get("sz")
        if size and size.isdigit():
            return int(size) / 100
    return default


def geometry(sp: ET.Element) -> tuple[int, int, int, int]:
    xfrm = sp.find("p:spPr/a:xfrm", NS)
    if xfrm is None:
        return 0, 0, 0, 0
    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return 0, 0, 0, 0
    return (
        int(off.get("x", "0")),
        int(off.get("y", "0")),
        int(ext.get("cx", "0")),
        int(ext.get("cy", "0")),
    )


def extract_text_shapes(xml_bytes: bytes, *, default_font_pt: float | None = None) -> tuple[list[ShapeSummary], int]:
    root = ET.fromstring(xml_bytes)
    text_shapes: list[ShapeSummary] = []
    for sp in root.findall(".//p:sp", NS):
        text = text_of_shape(sp)
        if not text:
            continue
        nv = sp.find("p:nvSpPr/p:cNvPr", NS)
        x, y, cx, cy = geometry(sp)
        text_shapes.append(
            ShapeSummary(
                shape_id=nv.get("id", "") if nv is not None else "",
                name=nv.get("name", "") if nv is not None else "",
                x=x,
                y=y,
                cx=cx,
                cy=cy,
                text=text,
                font_pt=first_font_pt(sp, default=default_font_pt),
            )
        )
    return sorted(text_shapes, key=lambda s: (s.y, s.x)), len(root.findall(".//p:sp", NS))


def slide_title(text_shapes: list[ShapeSummary]) -> str | None:
    if not text_shapes:
        return None
    sorted_shapes = sorted(
        text_shapes,
        key=lambda s: (
            s.y,
            -(s.font_pt or 0),
            -len(s.text),
        ),
    )
    return sorted_shapes[0].text if sorted_shapes[0].text else None


def summarize_slide(xml_bytes: bytes, slide_number: int, *, default_font_pt: float | None = None) -> SlideSummary:
    text_shapes, shape_count = extract_text_shapes(xml_bytes, default_font_pt=default_font_pt)
    return SlideSummary(
        slide_number=slide_number,
        title=slide_title(text_shapes),
        shape_count=shape_count,
        text_shapes=text_shapes,
    )


def load_slide_summaries(pptx: Path, *, slide_number: int | None = None, default_font_pt: float | None = None) -> list[SlideSummary]:
    with ZipFile(pptx) as zf:
        slides = []
        for num, part in iter_slide_parts(zf):
            if slide_number and num != slide_number:
                continue
            slides.append(summarize_slide(zf.read(part), num, default_font_pt=default_font_pt))
    return slides


def load_presentation_size(pptx: Path) -> PresentationSize:
    with ZipFile(pptx) as zf:
        root = ET.fromstring(zf.read("ppt/presentation.xml"))
    node = root.find("p:sldSz", NS)
    if node is None:
        return PresentationSize(cx=0, cy=0)
    return PresentationSize(
        cx=int(node.get("cx", "0")),
        cy=int(node.get("cy", "0")),
    )


def slide_page_position(pptx: Path, slide_number: int) -> int:
    pkg_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    with ZipFile(pptx) as zf:
        pres = ET.fromstring(zf.read("ppt/presentation.xml"))
        rels = ET.fromstring(zf.read("ppt/_rels/presentation.xml.rels"))
    target_by_rel_id = {
        rel.get("Id"): rel.get("Target")
        for rel in rels.findall(f"{{{pkg_ns}}}Relationship")
    }
    ordered_targets: list[str] = []
    sld_id_list = pres.find(f"{{{NS['p']}}}sldIdLst")
    if sld_id_list is None:
        raise ValueError("presentation missing slide list")
    for node in sld_id_list.findall(f"{{{NS['p']}}}sldId"):
        rel_id = node.get(f"{{{NS['r']}}}id")
        target = target_by_rel_id.get(rel_id or "")
        if target:
            ordered_targets.append(target)
    target = f"slides/slide{slide_number}.xml"
    if target not in ordered_targets:
        raise ValueError(f"slide {slide_number} not found in presentation order")
    return ordered_targets.index(target) + 1
