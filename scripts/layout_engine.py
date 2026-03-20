#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from dataclasses import asdict, dataclass


@dataclass
class Box:
    x: int
    y: int
    cx: int
    cy: int


def grid_2x2(
    width: int,
    height: int,
    *,
    left: int,
    right: int,
    top: int,
    bottom: int,
    gap_x: int,
    gap_y: int,
) -> list[Box]:
    inner_w = width - left - right
    inner_h = height - top - bottom
    card_w = (inner_w - gap_x) // 2
    card_h = (inner_h - gap_y) // 2
    return [
        Box(left, top, card_w, card_h),
        Box(left + card_w + gap_x, top, card_w, card_h),
        Box(left, top + card_h + gap_y, card_w, card_h),
        Box(left + card_w + gap_x, top + card_h + gap_y, card_w, card_h),
    ]


def banner_with_callout(
    width: int,
    *,
    left: int,
    right: int,
    top: int,
    height: int,
    callout_w: int,
    gap: int,
) -> dict[str, Box]:
    body_w = width - left - right - callout_w - gap
    return {
        "body": Box(left, top, body_w, height),
        "callout": Box(left + body_w + gap, top + height // 2 - 190500, callout_w, 381000),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Compute reusable slide layout boxes.")
    parser.add_argument("--width", type=int, required=True)
    parser.add_argument("--height", type=int, required=True)
    args = parser.parse_args()
    payload = {
        "grid_2x2": [asdict(item) for item in grid_2x2(args.width, args.height, left=285750, right=285750, top=3236002, bottom=550000, gap_x=260000, gap_y=250000)],
        "banner": {key: asdict(value) for key, value in banner_with_callout(args.width, left=1300162, right=285750, top=1590000, height=850000, callout_w=2700000, gap=250000).items()},
    }
    print(json.dumps(payload, indent=2))


if __name__ == "__main__":
    main()
