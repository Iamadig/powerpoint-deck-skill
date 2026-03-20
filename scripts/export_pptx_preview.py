#!/usr/bin/env python3
from __future__ import annotations

import argparse
import fcntl
import shutil
import subprocess
import time
from pathlib import Path


def default_work_root() -> Path:
    return Path.cwd() / ".pptx-work"


def parse_slides(value: str | None) -> list[int]:
    if not value:
        return []
    slides: list[int] = []
    for chunk in value.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "-" in chunk:
            start, end = chunk.split("-", 1)
            slides.extend(range(int(start), int(end) + 1))
        else:
            slides.append(int(chunk))
    return sorted(set(slides))


def export_pdf_via_powerpoint(pptx: Path, pdf_out: Path, *, attempts: int = 3, settle_seconds: float = 1.0) -> Path:
    if shutil.which("osascript") is None:
        raise SystemExit("osascript not found; cannot drive PowerPoint for PDF export")
    if not Path("/Applications/Microsoft PowerPoint.app").exists():
        raise SystemExit("Microsoft PowerPoint.app not found; cannot export PPTX to PDF")

    script = f"""
set inFile to POSIX file "{pptx}"
set outFile to POSIX file "{pdf_out}"
tell application "Microsoft PowerPoint"
  activate
  open inFile
  save active presentation in outFile as save as PDF
  close active presentation saving no
end tell
"""
    last_error: Exception | None = None
    lock_path = default_work_root() / "powerpoint-export.lock"
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    with lock_path.open("w") as lock_file:
        fcntl.flock(lock_file.fileno(), fcntl.LOCK_EX)
        for attempt in range(1, attempts + 1):
            pdf_out.unlink(missing_ok=True)
            try:
                result = subprocess.run(
                    ["osascript", "-e", script],
                    check=True,
                    capture_output=True,
                    text=True,
                )
                _ = result.stdout
                time.sleep(settle_seconds)
                if pdf_out.exists() and pdf_out.stat().st_size > 0:
                    return pdf_out
                last_error = SystemExit(f"PowerPoint export did not create {pdf_out}")
            except subprocess.CalledProcessError as exc:  # pragma: no cover - platform-specific automation
                stderr = exc.stderr.strip() if exc.stderr else str(exc)
                last_error = SystemExit(f"PowerPoint export failed: {stderr}")
            except Exception as exc:  # pragma: no cover - platform-specific automation
                last_error = exc
            if attempt < attempts:
                time.sleep(1.0 * attempt)
    if isinstance(last_error, SystemExit):
        raise last_error
    raise SystemExit(f"PowerPoint export failed for {pptx}: {last_error}")


def rasterize_pdf(
    pdf_path: Path,
    output_dir: Path,
    *,
    slides: list[int] | None = None,
    dpi: int = 192,
    prefix: str = "slide",
) -> list[Path]:
    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm is None:
        raise SystemExit("pdftoppm not found; install poppler to rasterize PDF previews")
    output_dir.mkdir(parents=True, exist_ok=True)
    targets = slides or []
    rendered: list[Path] = []
    if not targets:
        base = output_dir / prefix
        subprocess.run([pdftoppm, "-png", "-r", str(dpi), str(pdf_path), str(base)], check=True)
        rendered.extend(sorted(output_dir.glob(f"{prefix}-*.png")))
        return rendered

    for slide in targets:
        base = output_dir / f"{prefix}-{slide:03d}"
        subprocess.run(
            [
                pdftoppm,
                "-png",
                "-r",
                str(dpi),
                "-f",
                str(slide),
                "-l",
                str(slide),
                str(pdf_path),
                str(base),
            ],
            check=True,
        )
        rendered.extend(sorted(output_dir.glob(f"{prefix}-{slide:03d}-*.png")))
    return rendered


def export_preview_images(
    pptx: Path,
    output_dir: Path | None = None,
    *,
    slides: list[int] | None = None,
    dpi: int = 192,
    keep_pdf: bool = False,
    pdf_out: Path | None = None,
) -> tuple[Path, list[Path]]:
    if output_dir is None:
        output_dir = default_work_root() / "previews" / pptx.stem
    cache_mode = pdf_out is None
    if pdf_out is None:
        pdf_out = default_work_root() / "cache" / f"{pptx.stem}.pdf"
    pdf_out.parent.mkdir(parents=True, exist_ok=True)
    if pdf_out.exists() and pdf_out.stat().st_mtime >= pptx.stat().st_mtime and pdf_out.stat().st_size > 0:
        pdf_path = pdf_out
    else:
        pdf_path = export_pdf_via_powerpoint(pptx, pdf_out)
    images = rasterize_pdf(pdf_path, output_dir, slides=slides, dpi=dpi)
    if not keep_pdf and not cache_mode:
        pdf_path.unlink(missing_ok=True)
    return pdf_path, images


def main() -> None:
    parser = argparse.ArgumentParser(description="Export PPTX slide previews via PowerPoint PDF export and pdftoppm rasterization.")
    parser.add_argument("--pptx", required=True)
    parser.add_argument("--output-dir")
    parser.add_argument("--slides", help="Comma list or ranges, e.g. 11 or 1,3,5-7")
    parser.add_argument("--dpi", type=int, default=192)
    parser.add_argument("--keep-pdf", action="store_true")
    parser.add_argument("--pdf-out")
    args = parser.parse_args()

    pptx = Path(args.pptx)
    output_dir = Path(args.output_dir) if args.output_dir else None
    slides = parse_slides(args.slides)
    pdf_out = Path(args.pdf_out) if args.pdf_out else None
    _, images = export_preview_images(
        pptx,
        output_dir,
        slides=slides,
        dpi=args.dpi,
        keep_pdf=args.keep_pdf,
        pdf_out=pdf_out,
    )
    for image in images:
        print(image)


if __name__ == "__main__":
    main()
