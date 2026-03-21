# PowerPoint Deck Skill

Keep it simple.

Use this skill to:

1. inspect a PowerPoint deck
2. generate new slides in the same design system
3. compare 2-3 strong variants
4. export one final PPTX

## Install

From the repo root:

```bash
pip install -e .
```

Then use:

```bash
pptx-skill --help
```

## Core flow

Inspect a deck:

```bash
pptx-skill inspect /path/to/deck.pptx --format json
```

Extract design system:

```bash
pptx-skill design-system /path/to/deck.pptx \
  --output /tmp/design_system.json
```

Write a `slide.json` that follows:

- `assets/slide_spec.schema.json`

Generate options first:

```bash
pptx-skill make /path/to/deck.pptx \
  --content-json /tmp/slide.json
```

This now:
- creates a variant manifest in `.pptx-work/variants/<deck>/variants.json`
- creates a board in `.pptx-work/variants/<deck>/board.html`
- recommends one option
- does not write the working draft yet

Choose one explicitly:

```bash
pptx-skill make /path/to/deck.pptx \
  --content-json /tmp/slide.json \
  --variant 1
```

Or accept the recommendation:

```bash
pptx-skill make /path/to/deck.pptx \
  --content-json /tmp/slide.json \
  --auto
```

Generate variants:

```bash
pptx-skill variants /path/to/deck.pptx \
  --content-json /tmp/slide.json \
  --output /tmp/variants.json
```

Preview a slide:

```bash
pptx-skill preview /path/to/deck.pptx --slide 12
```

Export final:

```bash
pptx-skill export /path/to/deck.pptx --name final-name
```
