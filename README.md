# powerpoint-deck

Codex skill for two jobs:

1. inspect an existing PowerPoint deck
2. add a new slide in the same design system

Public entrypoints:

- `scripts/inspect_pptx.py`
- `scripts/make_slide.py`
- `scripts/preview_slide.py`

## Quickstart

Extract deck system + content:

```bash
python3 scripts/inspect_pptx.py --pptx /path/to/deck.pptx --format json
```

Make a new slide:

```bash
python3 scripts/make_slide.py --pptx /path/to/deck.pptx --content-json /tmp/slide.json
```

Preview one slide:

```bash
python3 scripts/preview_slide.py --pptx /path/to/deck.pptx --slide 5
```

## Content JSON

```json
{
  "title": "Who Feels This First",
  "section_label": "09 / Customer Segments",
  "hide_banner": true,
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
