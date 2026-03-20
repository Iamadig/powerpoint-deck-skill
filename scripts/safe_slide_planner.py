#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from dataclasses import asdict, dataclass
from pathlib import Path

from classify_slide_archetypes import classify
from generate_text_budgets import budget_for_shape, role_guess, title_id
from pptx_common import SlideSummary, load_slide_summaries


@dataclass
class ContentCard:
    eyebrow: str = ""
    title: str = ""
    body: str = ""
    meta: str = ""


@dataclass
class ContentSpec:
    section_label: str = ""
    title: str = ""
    subtitle: str = ""
    lead: str = ""
    footer: str = ""
    hide_footer: bool = False
    cards: list[ContentCard] | None = None


@dataclass
class SlotMatch:
    needed: int
    matched: int
    capacities: list[int]
    requirements: list[int]


@dataclass
class SlidePlan:
    slide_number: int
    title: str | None
    archetype: str
    score: int
    verdict: str
    matches: dict[str, SlotMatch]
    warnings: list[str]
    notes: list[str]


@dataclass
class ShapeSlot:
    shape_id: str
    shape_name: str
    role: str
    safe_chars_est: int
    x: int
    y: int


def load_content_spec(path: Path) -> ContentSpec:
    payload = json.loads(path.read_text())
    cards_payload = payload.get("cards") or payload.get("regions") or []
    cards = [
        ContentCard(
            eyebrow=item.get("eyebrow", ""),
            title=item.get("title", ""),
            body=item.get("body", ""),
            meta=item.get("meta", ""),
        )
        for item in cards_payload
    ]
    return ContentSpec(
        section_label=payload.get("section_label", ""),
        title=payload.get("title", ""),
        subtitle=payload.get("subtitle", ""),
        lead=payload.get("lead", ""),
        footer=payload.get("footer", ""),
        hide_footer=bool(payload.get("hide_footer", False)),
        cards=cards,
    )


def slide_budgets(slide: SlideSummary) -> list[dict[str, object]]:
    slide_title_id = title_id(slide.text_shapes)
    return [budget_for_shape(shape, role_guess(shape, slide_title_id)) for shape in slide.text_shapes]


def slide_slots(slide: SlideSummary) -> list[ShapeSlot]:
    title_shape_id = title_id(slide.text_shapes)
    wide_non_footer = sorted(
        [
            shape
            for shape in slide.text_shapes
            if shape.shape_id != title_shape_id
            and not (shape.y < 700000 and (shape.font_pt or 0) <= 12)
            and shape.width_pt >= 400
            and shape.y < 5800000
        ],
        key=lambda item: (item.y, item.x),
    )
    subtitle_shape_id = wide_non_footer[0].shape_id if wide_non_footer else None
    lead_shape_id = wide_non_footer[1].shape_id if len(wide_non_footer) > 1 else None
    budgets_by_id = {
        str(item["shape_id"]): item
        for item in slide_budgets(slide)
    }
    slots: list[ShapeSlot] = []
    for shape in slide.text_shapes:
        budget = budgets_by_id[shape.shape_id]
        font = shape.font_pt or 0
        width = shape.width_pt
        if title_shape_id and shape.shape_id == title_shape_id:
            role = "title"
        elif shape.y < 700000 and font <= 12:
            role = "section-label"
        elif subtitle_shape_id and shape.shape_id == subtitle_shape_id:
            role = "subtitle"
        elif lead_shape_id and shape.shape_id == lead_shape_id:
            role = "lead"
        elif budget["role"] == "footer" or (shape.y > 5800000 and width >= 300):
            role = "footer"
        elif font <= 10.5 and 60 <= width < 400 and 5000000 <= shape.y <= 5900000:
            role = "card-meta"
        elif font >= 13 and 60 <= width < 400 and 1800000 <= shape.y < 5600000:
            role = "card-title"
        elif 60 <= width < 400 and 2200000 <= shape.y < 5900000:
            role = "card-body"
        else:
            role = "body"
        slots.append(
            ShapeSlot(
                shape_id=shape.shape_id,
                shape_name=shape.name,
                role=role,
                safe_chars_est=int(budget["safe_chars_est"]),
                x=shape.x,
                y=shape.y,
            )
        )
    return slots


def flatten_card_text(cards: list[ContentCard]) -> list[str]:
    items: list[str] = []
    for card in cards:
        if card.eyebrow:
            items.append(card.eyebrow)
        if card.title:
            items.append(card.title)
    return items


def flatten_card_bodies(cards: list[ContentCard]) -> list[str]:
    return [card.body for card in cards if card.body]


def flatten_card_meta(cards: list[ContentCard]) -> list[str]:
    return [card.meta for card in cards if card.meta]


def sort_slots(slots: list[ShapeSlot]) -> list[ShapeSlot]:
    return sorted(slots, key=lambda item: (item.safe_chars_est, item.y, item.x))


def role_slots(slots: list[ShapeSlot], roles: set[str]) -> list[ShapeSlot]:
    return [slot for slot in slots if slot.role in roles]


def assign_single_with_preferences(
    available: list[ShapeSlot],
    text: str,
    preferred_role_groups: list[set[str]],
) -> list[ShapeSlot]:
    if not text:
        return []
    for roles in preferred_role_groups:
        candidates = sort_slots(role_slots(available, roles))
        for slot in candidates:
            if slot.safe_chars_est >= len(text):
                return [slot]
    return []


def assign_many_with_preferences(
    available: list[ShapeSlot],
    texts: list[str],
    preferred_role_groups: list[set[str]],
) -> dict[int, ShapeSlot]:
    if not texts:
        return {}
    remaining = list(available)
    assignments: dict[int, ShapeSlot] = {}
    pending = sorted(
        [(index, text) for index, text in enumerate(texts)],
        key=lambda item: len(item[1]),
        reverse=True,
    )
    for roles in preferred_role_groups:
        if not pending:
            break
        candidates = sort_slots(role_slots(remaining, roles))
        next_pending: list[tuple[int, str]] = []
        for index, text in pending:
            for slot_index, slot in enumerate(candidates):
                if slot.safe_chars_est >= len(text):
                    assignments[index] = slot
                    remaining.remove(slot)
                    del candidates[slot_index]
                    break
            else:
                next_pending.append((index, text))
        pending = next_pending
    return assignments


def assignment_counts(slide: SlideSummary, spec: ContentSpec) -> dict[str, SlotMatch]:
    cards = spec.cards or []
    slots = slide_slots(slide)
    unused = list(slots)
    matches: dict[str, SlotMatch] = {}

    def capacities(groups: list[set[str]]) -> list[int]:
        caps: list[int] = []
        seen: set[str] = set()
        for roles in groups:
            for slot in sort_slots(role_slots(slots, roles)):
                if slot.shape_id in seen:
                    continue
                caps.append(slot.safe_chars_est)
                seen.add(slot.shape_id)
        return sorted(caps)

    def remove_assigned(items: list[ShapeSlot]) -> None:
        assigned_ids = {item.shape_id for item in items}
        unused[:] = [slot for slot in unused if slot.shape_id not in assigned_ids]

    section_assigned = assign_single_with_preferences(unused, spec.section_label, [{"section-label"}])
    matches["section_label"] = SlotMatch(
        needed=1 if spec.section_label else 0,
        matched=len(section_assigned),
        capacities=capacities([{"section-label"}]),
        requirements=[len(spec.section_label)] if spec.section_label else [],
    )
    remove_assigned(section_assigned)

    title_assigned = assign_single_with_preferences(unused, spec.title, [{"title"}])
    matches["title"] = SlotMatch(
        needed=1 if spec.title else 0,
        matched=len(title_assigned),
        capacities=capacities([{"title"}]),
        requirements=[len(spec.title)] if spec.title else [],
    )
    remove_assigned(title_assigned)

    subtitle_assigned = assign_single_with_preferences(unused, spec.subtitle, [{"subtitle"}])
    matches["subtitle"] = SlotMatch(
        needed=1 if spec.subtitle else 0,
        matched=len(subtitle_assigned),
        capacities=capacities([{"subtitle"}]),
        requirements=[len(spec.subtitle)] if spec.subtitle else [],
    )
    remove_assigned(subtitle_assigned)

    lead_assigned = assign_single_with_preferences(unused, spec.lead, [{"lead"}, {"body"}])
    matches["lead"] = SlotMatch(
        needed=1 if spec.lead else 0,
        matched=len(lead_assigned),
        capacities=capacities([{"lead"}, {"body"}]),
        requirements=[len(spec.lead)] if spec.lead else [],
    )
    remove_assigned(lead_assigned)

    card_texts = flatten_card_text(cards)
    card_text_assigned = assign_many_with_preferences(unused, card_texts, [{"card-title"}])
    matches["card_text"] = SlotMatch(
        needed=len(card_texts),
        matched=len(card_text_assigned),
        capacities=capacities([{"card-title"}]),
        requirements=sorted([len(item) for item in card_texts]),
    )
    remove_assigned(list(card_text_assigned.values()))

    card_bodies = flatten_card_bodies(cards)
    card_body_assigned = assign_many_with_preferences(unused, card_bodies, [{"card-body"}])
    matches["card_body"] = SlotMatch(
        needed=len(card_bodies),
        matched=len(card_body_assigned),
        capacities=capacities([{"card-body"}]),
        requirements=sorted([len(item) for item in card_bodies]),
    )
    remove_assigned(list(card_body_assigned.values()))

    card_meta = flatten_card_meta(cards)
    card_meta_assigned = assign_many_with_preferences(unused, card_meta, [{"card-meta"}])
    matches["card_meta"] = SlotMatch(
        needed=len(card_meta),
        matched=len(card_meta_assigned),
        capacities=capacities([{"card-meta"}]),
        requirements=sorted([len(item) for item in card_meta]),
    )
    remove_assigned(list(card_meta_assigned.values()))

    footer_assigned = assign_single_with_preferences(unused, spec.footer, [{"footer"}, {"body"}])
    matches["footer"] = SlotMatch(
        needed=1 if spec.footer else 0,
        matched=len(footer_assigned),
        capacities=capacities([{"footer"}, {"body"}]),
        requirements=[len(spec.footer)] if spec.footer else [],
    )
    remove_assigned(footer_assigned)
    return matches


def fit_verdict(matches: dict[str, SlotMatch]) -> str:
    total_needed = sum(item.needed for item in matches.values())
    total_matched = sum(item.matched for item in matches.values())
    if total_needed == 0:
        return "unknown"
    if total_matched == total_needed:
        return "safe"
    if total_matched >= max(total_needed - 1, 1):
        return "partial"
    return "unsafe"


def score_slide(archetype: str, matches: dict[str, SlotMatch], card_count: int) -> tuple[int, list[str]]:
    score = 0
    notes: list[str] = []
    total_needed = sum(item.needed for item in matches.values())
    total_matched = sum(item.matched for item in matches.values())
    score += total_matched * 10
    score -= (total_needed - total_matched) * 18

    if card_count >= 4 and archetype == "multi-column-grid-with-footer":
        score += 14
        notes.append("4-card work tends to fit multi-column grid best")
    elif card_count >= 4 and archetype == "two-column-grid-with-banner":
        score -= 8
        notes.append("banner layouts often underfit 4 card-heavy blocks")
    elif card_count >= 4 and archetype == "two-column-comparison":
        score -= 16
        notes.append("comparison layouts rarely fit dense persona cards")

    if card_count == 0 and archetype == "cover-or-closing":
        score += 8
    return score, notes


def warnings_for(matches: dict[str, SlotMatch]) -> list[str]:
    warnings: list[str] = []
    labels = {
        "section_label": "section label",
        "title": "slide title",
        "subtitle": "subtitle",
        "lead": "lead text",
        "card_text": "card heading/eyebrow slots",
        "card_body": "card body slots",
        "card_meta": "card meta slots",
        "footer": "footer",
    }
    for key, item in matches.items():
        if item.needed and item.matched < item.needed:
            warnings.append(
                f"needs {item.needed} {labels[key]}; only {item.matched} fit current box budgets"
            )
    return warnings


def plan_slide(slide: SlideSummary, spec: ContentSpec) -> SlidePlan:
    archetype_summary = classify(slide)
    cards = spec.cards or []
    matches = assignment_counts(slide, spec)
    score, notes = score_slide(archetype_summary.archetype, matches, len(cards))
    warnings = warnings_for(matches)
    return SlidePlan(
        slide_number=slide.slide_number,
        title=slide.title,
        archetype=archetype_summary.archetype,
        score=score,
        verdict=fit_verdict(matches),
        matches=matches,
        warnings=warnings,
        notes=notes,
    )


def render_match(name: str, match: SlotMatch) -> str:
    if match.needed == 0:
        return f"  - {name}: not needed"
    return (
        f"  - {name}: {match.matched}/{match.needed} fit "
        f"(req={match.requirements}, caps={match.capacities[:8]})"
    )


def render_plan(plan: SlidePlan) -> str:
    lines = [
        f"slide {plan.slide_number}: {plan.title or '(untitled)'}",
        f"  archetype: {plan.archetype}",
        f"  verdict: {plan.verdict}",
        f"  score: {plan.score}",
    ]
    for key in ("section_label", "title", "subtitle", "lead", "card_text", "card_body", "card_meta", "footer"):
        lines.append(render_match(key, plan.matches[key]))
    if plan.notes:
        lines.append(f"  notes: {'; '.join(plan.notes)}")
    if plan.warnings:
        lines.append(f"  warnings: {'; '.join(plan.warnings)}")
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description="Pick the safest existing slide archetype for new PPTX content.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--content-json", required=True, help="JSON file with title/subtitle/cards/footer.")
    parser.add_argument("--slide", type=int, help="Only score one slide.")
    parser.add_argument("--top", type=int, default=5)
    parser.add_argument("--format", choices=["text", "json"], default="text")
    args = parser.parse_args()

    spec = load_content_spec(Path(args.content_json))
    slides = load_slide_summaries(Path(args.pptx), slide_number=args.slide, default_font_pt=18.0)
    plans = sorted(
        (plan_slide(slide, spec) for slide in slides),
        key=lambda item: (item.score, item.verdict == "safe"),
        reverse=True,
    )
    plans = plans[: max(args.top, 1)]

    if args.format == "json":
        print(json.dumps([asdict(item) for item in plans], indent=2))
        return

    safe = next((item for item in plans if item.verdict == "safe"), None)
    if safe:
        print(f"recommended: slide {safe.slide_number} ({safe.archetype})")
    else:
        print("recommended: no safe existing archetype; redesign or shorten copy")
    for plan in plans:
        print()
        print(render_plan(plan))


if __name__ == "__main__":
    main()
