#!/usr/bin/env python3
"""Deterministically reduce oversized Matplotlib scatter collections for PPTX.

The optimizer only removes repeated ``<use>`` markers inside ``PathCollection``
groups. Text, axes, contours, hulls, centroids, annotations, and all non-scatter
geometry remain unchanged. Source SVG files are never modified in place.
"""

from __future__ import annotations

import argparse
import json
import math
import shutil
from pathlib import Path

from lxml import etree


def allocate_targets(sizes: list[int], budget: int, minimum: int) -> list[int]:
    if sum(sizes) <= budget:
        return sizes[:]
    base = [min(size, minimum) for size in sizes]
    if sum(base) > budget:
        scale = budget / sum(base)
        targets = [max(1, min(size, math.floor(value * scale))) for size, value in zip(sizes, base)]
    else:
        targets = base[:]
        remaining = budget - sum(targets)
        capacities = [size - target for size, target in zip(sizes, targets)]
        capacity_total = sum(capacities)
        if remaining > 0 and capacity_total > 0:
            raw = [remaining * capacity / capacity_total for capacity in capacities]
            additions = [min(capacity, math.floor(value)) for capacity, value in zip(capacities, raw)]
            for index, addition in enumerate(additions):
                targets[index] += addition
            residual = budget - sum(targets)
            order = sorted(
                range(len(sizes)),
                key=lambda index: (raw[index] - math.floor(raw[index]), capacities[index]),
                reverse=True,
            )
            for index in order:
                if residual <= 0:
                    break
                if targets[index] < sizes[index]:
                    targets[index] += 1
                    residual -= 1
    return targets


def evenly_spaced_indices(size: int, target: int) -> set[int]:
    if target >= size:
        return set(range(size))
    if target <= 1:
        return {size // 2}
    return {round(index * (size - 1) / (target - 1)) for index in range(target)}


def optimize_svg(source: Path, destination: Path, max_markers: int, minimum: int) -> dict:
    parser = etree.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False)
    tree = etree.parse(str(source), parser)
    root = tree.getroot()
    collections = root.xpath('//*[local-name()="g" and starts-with(@id,"PathCollection_")]')
    eligible = []
    fixed_markers = 0
    for collection in collections:
        uses = collection.xpath('.//*[local-name()="use"]')
        if len(uses) <= 1:
            fixed_markers += len(uses)
        else:
            eligible.append((collection, uses))
    sizes = [len(uses) for _, uses in eligible]
    budget = max(0, max_markers - fixed_markers)
    targets = allocate_targets(sizes, budget, minimum) if sizes else []
    removed = 0
    details = []
    for (collection, uses), target in zip(eligible, targets):
        keep = evenly_spaced_indices(len(uses), target)
        for index, use in enumerate(uses):
            if index not in keep:
                use.getparent().remove(use)
                removed += 1
        details.append({"id": collection.get("id"), "before": len(uses), "after": len(keep)})
    destination.parent.mkdir(parents=True, exist_ok=True)
    tree.write(str(destination), encoding="utf-8", xml_declaration=True)
    return {
        "file": source.name,
        "markers_before": fixed_markers + sum(sizes),
        "markers_after": fixed_markers + sum(targets),
        "removed": removed,
        "collections": details,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--max-markers-per-panel", type=int, default=900)
    parser.add_argument("--minimum-per-collection", type=int, default=20)
    args = parser.parse_args()
    if args.input_dir.resolve() == args.output_dir.resolve():
        raise ValueError("Output directory must differ from input directory")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    reports = []
    for source in sorted(args.input_dir.glob("panel-*.svg")):
        reports.append(
            optimize_svg(
                source,
                args.output_dir / source.name,
                args.max_markers_per_panel,
                args.minimum_per_collection,
            )
        )
    for extra in ("manifest.json", "shared-legend.json"):
        source = args.input_dir / extra
        if source.exists():
            shutil.copy2(source, args.output_dir / extra)
    report_path = args.output_dir / "optimization-report.json"
    report_path.write_text(json.dumps(reports, indent=2), encoding="utf-8")
    print(f"Optimized {len(reports)} SVG panels into {args.output_dir}")
    print(f"Report: {report_path}")


if __name__ == "__main__":
    main()
