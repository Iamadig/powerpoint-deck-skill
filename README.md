# PowerPoint Deck Skill

Keep it simple.

Use this skill to:

1. inspect a PowerPoint deck
2. generate new slides in the same design system
3. compare 2-3 strong variants
4. export one final PPTX

## Core flow

Extract design system:

```bash
python3 scripts/extract_design_system.py \
  --pptx /path/to/deck.pptx \
  --output /tmp/design_system.json
```

Write a `slide.json` that follows:

- `assets/slide_spec.schema.json`

Generate one slide:

```bash
python3 scripts/make_slide.py \
  --pptx /path/to/deck.pptx \
  --content-json /tmp/slide.json
```

Generate variants:

```bash
python3 scripts/generate_variants.py \
  --pptx /path/to/deck.pptx \
  --content-json /tmp/slide.json \
  --output /tmp/variants.json
```

Build variant board:

```bash
python3 scripts/variant_board.py \
  --manifest /tmp/variants.json \
  --output /tmp/variant-board.html
```

Export final:

```bash
python3 scripts/export_final.py \
  --pptx /path/to/deck.pptx \
  --name final-name
```
