#!/usr/bin/env python3
"""Render a PPTX with Microsoft PowerPoint on macOS for authoritative visual QA."""

from __future__ import annotations

import argparse
import shutil
import subprocess
import tempfile
from pathlib import Path


def apple_string(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def find_pdftoppm() -> str:
    found = shutil.which("pdftoppm")
    if found:
        return found
    root = Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies"
    candidates = [
        root / "bin/pdftoppm",
        root / "bin/override/pdftoppm",
        root / "native/poppler/bin/pdftoppm",
        root / "native/poppler/poppler/bin/pdftoppm",
    ]
    for bundled in candidates:
        if bundled.exists():
            return str(bundled)
    raise FileNotFoundError("pdftoppm was not found; install Poppler or pass a runtime with it")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--pdf", type=Path, required=True)
    parser.add_argument("--png", type=Path, required=True)
    parser.add_argument("--dpi", type=int, default=100)
    args = parser.parse_args()
    if shutil.which("osascript") is None:
        raise SystemExit("This renderer requires macOS and Microsoft PowerPoint")
    with tempfile.TemporaryDirectory(prefix="pptx-powerpoint-", dir="/private/tmp") as temp_dir:
        temp_dir = Path(temp_dir)
        temp_pptx = temp_dir / "input.pptx"
        temp_pdf = temp_dir / "output.pdf"
        temp_png_prefix = temp_dir / "slide"
        shutil.copy2(args.pptx, temp_pptx)
        script = [
            f'set pptFile to POSIX file "{apple_string(str(temp_pptx))}"',
            f'set pdfFile to POSIX file "{apple_string(str(temp_pdf))}"',
            "with timeout of 600 seconds",
            'tell application "Microsoft PowerPoint"',
            "open pptFile",
            "set p to active presentation",
            "save p in pdfFile as save as PDF",
            "close p saving no",
            "end tell",
            "end timeout",
        ]
        command = ["osascript"]
        for line in script:
            command.extend(["-e", line])
        subprocess.run(command, check=True)
        subprocess.run(
            [find_pdftoppm(), "-png", "-r", str(args.dpi), "-singlefile", str(temp_pdf), str(temp_png_prefix)],
            check=True,
        )
        args.pdf.parent.mkdir(parents=True, exist_ok=True)
        args.png.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(temp_pdf, args.pdf)
        shutil.copy2(temp_png_prefix.with_suffix(".png"), args.png)
    print(f"PowerPoint PDF: {args.pdf}")
    print(f"PowerPoint PNG: {args.png}")


if __name__ == "__main__":
    main()
