#!/usr/bin/env python3
"""Validate a directory of editable, code-generated SVG figure panels."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from lxml import etree


PANEL_RE = re.compile(r"panel-(\d{2})\.svg$")
TITLE_RE = re.compile(r"^\s*([a-z])[\.\)]\s+\S", re.IGNORECASE)
XLINK_HREF = "{http://www.w3.org/1999/xlink}href"


def local_name(element: etree._Element) -> str:
    if not isinstance(element.tag, str):
        return ""
    return etree.QName(element).localname


def parse_view_box(root: etree._Element) -> tuple[float, float, float, float] | None:
    raw = root.get("viewBox")
    if not raw:
        return None
    try:
        values = tuple(float(value) for value in re.split(r"[\s,]+", raw.strip()))
    except ValueError:
        return None
    if len(values) != 4 or values[2] <= 0 or values[3] <= 0:
        return None
    return values


def validate_directory(svg_dir: Path, manifest_path: Path | None) -> list[str]:
    errors: list[str] = []
    paths = sorted(svg_dir.glob("panel-*.svg"))
    if not 2 <= len(paths) <= 12:
        errors.append(f"Expected 2–12 panel SVGs; found {len(paths)}")

    expected_names = [f"panel-{index:02d}.svg" for index in range(1, len(paths) + 1)]
    actual_names = [path.name for path in paths]
    if actual_names != expected_names:
        errors.append(
            "Panel filenames must be contiguous and ordered: "
            f"expected {expected_names}, found {actual_names}"
        )

    parser = etree.XMLParser(resolve_entities=False, no_network=True, recover=False)
    reports: list[dict[str, object]] = []
    for index, path in enumerate(paths, start=1):
        match = PANEL_RE.fullmatch(path.name)
        if not match or int(match.group(1)) != index:
            continue
        try:
            root = etree.parse(str(path), parser).getroot()
        except (OSError, etree.XMLSyntaxError) as exc:
            errors.append(f"{path.name}: invalid XML ({exc})")
            continue

        view_box = parse_view_box(root)
        if view_box is None:
            errors.append(f"{path.name}: missing or invalid positive viewBox")

        images = [element for element in root.iter() if local_name(element) == "image"]
        if images:
            errors.append(f"{path.name}: contains {len(images)} raster <image> element(s)")

        external_refs: list[str] = []
        for element in root.iter():
            for key in ("href", XLINK_HREF):
                href = element.get(key)
                if href and (
                    href.startswith("data:image")
                    or href.startswith("http://")
                    or href.startswith("https://")
                    or href.startswith("file:")
                ):
                    external_refs.append(href[:80])
        if external_refs:
            errors.append(f"{path.name}: contains raster or external references")

        texts = [
            "".join(element.itertext()).strip()
            for element in root.iter()
            if local_name(element) == "text"
        ]
        texts = [text for text in texts if text]
        if not texts:
            errors.append(f"{path.name}: contains no editable <text> elements")

        expected_letter = chr(ord("a") + index - 1)
        title_matches = [
            text
            for text in texts
            if (title_match := TITLE_RE.match(text))
            and title_match.group(1).lower() == expected_letter
        ]
        if len(title_matches) != 1:
            errors.append(
                f"{path.name}: expected exactly one '{expected_letter}.' panel title; "
                f"found {len(title_matches)}"
            )

        reports.append(
            {
                "file": path.name,
                "viewBox": list(view_box) if view_box else None,
                "text_count": len(texts),
                "image_count": len(images),
                "title": title_matches[0] if len(title_matches) == 1 else None,
            }
        )

    if manifest_path:
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            errors.append(f"Manifest is unreadable: {exc}")
        else:
            if manifest.get("panel_count") != len(paths):
                errors.append(
                    "manifest.panel_count does not match the number of SVG panels"
                )
            manifest_files = [
                str(panel.get("file", "")) for panel in manifest.get("panels", [])
            ]
            if manifest_files != actual_names:
                errors.append(
                    "manifest.panels file order does not match filename order"
                )

    print(
        json.dumps(
            {
                "svg_dir": str(svg_dir),
                "panel_count": len(paths),
                "panels": reports,
                "errors": errors,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return errors


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("svg_dir", type=Path)
    parser.add_argument("--manifest", type=Path)
    args = parser.parse_args()
    if not args.svg_dir.is_dir():
        parser.error(f"SVG directory does not exist: {args.svg_dir}")
    manifest = args.manifest
    if manifest is None and (args.svg_dir / "manifest.json").exists():
        manifest = args.svg_dir / "manifest.json"
    errors = validate_directory(args.svg_dir, manifest)
    raise SystemExit(1 if errors else 0)


if __name__ == "__main__":
    main()
