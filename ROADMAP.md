# Roadmap

## Principle

Keep this small.
Do not build a Figma clone.
Do not build a full slide editor.

Goal:
- inspect deck
- generate good new slides
- preview a few strong options
- export one final PPTX

## Current state

Works now:
- deck inspection
- safe native archetype reuse
- one mature synthesized archetype: 4-card
- working draft flow
- preview current draft
- explicit final export
- cleanup

Weak now:
- arbitrary freeform layout invention
- more than a few slide archetypes
- multi-variant selection
- true design-system extraction beyond rough heuristics

## Next milestone: Variant board

Smallest useful addition.

Build:
- `generate_variants.py`
- `variant_board.py`

Behavior:
- same content
- 2–3 layout variants max
- render previews side by side
- choose one

Not in scope:
- drag and drop
- direct manipulation
- persistent canvas state

## Why variant board first

This gets most of the value of a canvas without the weight:
- visual comparison
- better design iteration
- lower risk than full editor UI
- keeps the workflow fast

## Path to arbitrary layout invention

Do not extend the current hand-coded archetype approach forever.
Split the system into four parts.

### 1. Extractor

Keep Python.
Read existing PPTX and emit:
- typography tokens
- colors
- spacing rhythm
- card/container styles
- common regions

Output:
- `design_system.json`

### 2. Planner

Codex creates a semantic slide spec.

Output:
- `slide_spec.json`

### 3. Layout engine

Use an actual layout solver instead of hardcoding positions.

Best starting point:
- `Yoga`

Possible later addition:
- `Kiwi`

### 4. Renderer

Move long-term rendering to:
- `PptxGenJS`

Keep Python where it still helps:
- extraction
- inspection

## Build order

### Phase 1

- ship variant board
- keep outputs in one working draft flow

### Phase 2

- add `design_system.json` extractor
- add `slide_spec.json`

### Phase 3

- Yoga proof of concept
- generate 2–3 freeform layouts from one slide spec

### Phase 4

- evaluate `PptxGenJS` renderer migration
- stop hand-editing PPT XML for new layouts
