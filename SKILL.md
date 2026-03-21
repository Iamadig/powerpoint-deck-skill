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

Install once if needed:

```bash
pip install -e .
```

### Extract deck system + content

```bash
pptx-skill inspect <deck.pptx> \
  --format json
```

This returns:
- slide size
- main colors
- font sizes
- slide archetypes
- per-slide text content + geometry

Clean design-system JSON:

```bash
pptx-skill design-system <deck.pptx> \
  --output /tmp/design_system.json
```

### Make one new slide

```bash
pptx-skill make <deck.pptx> \
  --content-json <slide.json>
```

This:
- generates options first
- writes a variant manifest at `.pptx-work/variants/<deck>/variants.json`
- writes a board at `.pptx-work/variants/<deck>/board.html`
- recommends one variant
- does not update the working draft yet

Optional:

```bash
pptx-skill make <deck.pptx> \
  --content-json <slide.json> \
  --variant 1
```

This applies the chosen variant to the working draft.

Or accept the recommendation:

```bash
pptx-skill make <deck.pptx> \
  --content-json <slide.json> \
  --auto
```

Named export still works:

```bash
pptx-skill make <deck.pptx> \
  --content-json <slide.json> \
  --auto \
  --save-as final-name
```

If the user gives a plain-English brief, Codex should:
1. inspect the deck
2. write a small `slide.json` that follows `assets/slide_spec.schema.json`
3. call `make_slide.py`

### Generate variants

```bash
pptx-skill variants <deck.pptx> \
  --content-json <slide.json> \
  --output /tmp/variants.json
```

This creates 2-3 safe variants from different native archetypes when possible.

### Variant board

```bash
pptx-skill preview <deck.pptx> --slide <n>
```

If a working draft exists, preview uses that by default.

### Export final

```bash
pptx-skill export <deck.pptx> --name final-name
```

### Clean artifacts

```bash
pptx-skill clean <deck.pptx>
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
- no silent auto-pick by default
- keep outputs repo-local in `.pptx-work/`
- one working draft; no version spam unless asked
- variant board, not full canvas editor
- Codex plans the content; scripts render it

## Internal Scripts

Most files in `scripts/` are internal helpers. Normal use should only need:
- `pptx-skill inspect`
- `pptx-skill design-system`
- `pptx-skill make`
- `pptx-skill variants`
- `pptx-skill preview`
- `pptx-skill export`
- `pptx-skill clean`
