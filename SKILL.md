---
name: powerpoint-deck
description: Extract the design system and content from an existing PowerPoint deck, then add new slides in the same style.
---

# PowerPoint Deck

Use this skill for exactly two things:

1. extract a deck cleanly
2. make one new slide in the same design system

Codex is the intelligence layer.
The scripts are deterministic.

## Commands

### Extract deck system + content

```bash
python3 scripts/inspect_pptx.py \
  --pptx <deck.pptx> \
  --format json
```

This returns:
- slide size
- main colors
- font sizes
- slide archetypes
- per-slide text content + geometry

### Make one new slide

```bash
python3 scripts/make_slide.py \
  --pptx <deck.pptx> \
  --content-json <slide.json>
```

This:
- finds a safe existing slide pattern when possible
- otherwise synthesizes a 4-card slide when the payload fits that shape
- writes the revised deck to `.pptx-work/out/`
- exports a preview PNG to `.pptx-work/previews/`

If the user gives a plain-English brief, Codex should:
1. inspect the deck
2. write a small `slide.json`
3. call `make_slide.py`

### Preview one slide

```bash
python3 scripts/preview_slide.py \
  --pptx <deck.pptx> \
  --slide <n>
```

## Content JSON

Minimal example:

```json
{
  "title": "Who Feels This First",
  "section_label": "09 / Customer Segments",
  "cards": [
    {
      "title": "Solo Developers",
      "quote": "I keep re-explaining the same project.",
      "pain": "Pain: context decays between sessions",
      "value": "Value: the next agent inherits context"
    },
    {
      "title": "Small Teams (2-5)",
      "quote": "We ship fast, but nobody knows why.",
      "pain": "Pain: coordination overhead",
      "value": "Value: handoffs + decision trails"
    },
    {
      "title": "Engineering Leads",
      "quote": "I need to understand why, not just what.",
      "pain": "Pain: invisible decisions",
      "value": "Value: auditable reasoning"
    },
    {
      "title": "Tool Builders",
      "quote": "I want memory without vendor lock-in.",
      "pain": "Pain: vendor lock-in",
      "value": "Value: git-native + protocol-level"
    }
  ]
}
```

## Rules

- content first
- reuse existing archetypes when safe
- no forced fit
- preview before trust
- keep outputs repo-local in `.pptx-work/`
- Codex plans the content; scripts render it

## Internal Scripts

Most files in `scripts/` are internal helpers. Normal use should only need:
- `inspect_pptx.py`
- `make_slide.py`
- `preview_slide.py`
