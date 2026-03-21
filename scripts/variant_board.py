#!/usr/bin/env python3
from __future__ import annotations

import argparse
import html
import json
from pathlib import Path


HTML_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Variant Board</title>
  <style>
    body {{ background:#0f1728; color:#f5f7fb; font-family:-apple-system,BlinkMacSystemFont,Segoe UI,sans-serif; margin:0; padding:32px; }}
    h1 {{ font-size:40px; margin:0 0 10px; }}
    p {{ color:#9ca7bb; margin:0 0 24px; }}
    .grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(420px,1fr)); gap:24px; }}
    .card {{ background:#182133; border:1px solid rgba(255,255,255,.1); border-radius:20px; padding:18px; }}
    .label {{ font-size:12px; letter-spacing:.12em; text-transform:uppercase; color:#21c6ff; margin-bottom:10px; }}
    .recommended {{ display:inline-block; margin-left:10px; padding:4px 8px; border-radius:999px; font-size:11px; letter-spacing:.08em; text-transform:uppercase; background:#14384a; color:#53d7ff; }}
    img {{ width:100%; height:auto; border-radius:12px; display:block; }}
    .meta {{ margin-top:12px; font-size:14px; color:#b7c2d6; }}
    code {{ color:#9de1ff; }}
  </style>
</head>
<body>
  <h1>Variant Board</h1>
  <p>{subtitle}</p>
  <div class="grid">
    {cards}
  </div>
</body>
</html>"""


CARD_TEMPLATE = """<div class="card">
  <div class="label">{label}{recommended}</div>
  <img src="{preview}" alt="{label}" />
  <div class="meta">
    archetype: <code>{archetype}</code><br/>
    source slide: <code>{source_slide}</code><br/>
    deck: <code>{deck}</code>
  </div>
</div>"""


def build_board(manifest: dict[str, object]) -> str:
    cards = []
    recommended = str(manifest.get("recommended_variant_id") or "")
    for variant in manifest["variants"]:
        cards.append(
            CARD_TEMPLATE.format(
                label=html.escape(str(variant["label"])),
                recommended='<span class="recommended">recommended</span>' if variant["id"] == recommended else "",
                preview=html.escape(str(variant["preview"])),
                archetype=html.escape(str(variant.get("archetype") or "custom")),
                source_slide=html.escape(str(variant.get("source_slide") or "-")),
                deck=html.escape(str(variant["deck"])),
            )
        )
    count = len(manifest["variants"])
    subtitle = "Pick one. Keep the workflow simple." if count > 1 else "Only one safe option fit this deck. Use it or shorten the content."
    return HTML_TEMPLATE.format(cards="\n".join(cards), subtitle=html.escape(subtitle))


def main() -> None:
    parser = argparse.ArgumentParser(description="Render a simple HTML board for slide variants.")
    parser.add_argument("--manifest", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    manifest = json.loads(Path(args.manifest).read_text())
    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(build_board(manifest))
    print(out)


if __name__ == "__main__":
    main()
