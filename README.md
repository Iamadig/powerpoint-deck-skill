# PowerPoint Deck Skill

Codex skill for two jobs:

1. inspect an existing PowerPoint deck
2. add a new slide in the same design system

The model handles content and structure.
The scripts handle rendering, fitting, preview, and export.

## Current workflow

Inspect:

```bash
python3 scripts/inspect_pptx.py --pptx /path/to/deck.pptx --format json
```

Iterate on one working draft:

```bash
python3 scripts/make_slide.py --pptx /path/to/deck.pptx --content-json /tmp/slide.json
```

Preview current draft:

```bash
python3 scripts/preview_slide.py --pptx /path/to/deck.pptx --slide 12
```

Export final only when approved:

```bash
python3 scripts/export_final.py --pptx /path/to/deck.pptx --name final-name
```

Clean artifacts:

```bash
python3 scripts/clean_workdir.py --pptx /path/to/deck.pptx
```

## Working files

- draft: `.pptx-work/out/<deck>/current.pptx`
- previews: `.pptx-work/previews/<deck>/current/`
- named exports: `.pptx-work/out/<deck>/<name>.pptx`

## Near-term direction

The next step is not a full editor.
It is a lightweight variant board:

- generate 2–3 slide variants
- render them side by side
- pick one
- export only the chosen version

That keeps the skill simple while improving design quality.
