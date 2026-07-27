#!/usr/bin/env python3
"""Compare a hand-adjusted PPTX with an automatically generated baseline."""

from __future__ import annotations

import argparse
import json
import re
from collections import Counter
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


EMU_PER_INCH = 914400
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}


def compose(parent: tuple[float, float, float, float], child: tuple[float, float, float, float]):
    psx, psy, ptx, pty = parent
    csx, csy, ctx, cty = child
    return psx * csx, psy * csy, psx * ctx + ptx, psy * cty + pty


def group_transform(group) -> tuple[float, float, float, float]:
    xfrm = group.find("./p:grpSpPr/a:xfrm", NS)
    if xfrm is None:
        return 1.0, 1.0, 0.0, 0.0
    off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
    child_off, child_ext = xfrm.find("a:chOff", NS), xfrm.find("a:chExt", NS)
    sx = int(ext.get("cx")) / max(1, int(child_ext.get("cx")))
    sy = int(ext.get("cy")) / max(1, int(child_ext.get("cy")))
    return (
        sx,
        sy,
        int(off.get("x")) - sx * int(child_off.get("x")),
        int(off.get("y")) - sy * int(child_off.get("y")),
    )


def shape_name(element) -> str:
    found = element.xpath(
        "./p:nvSpPr/p:cNvPr/@name | ./p:nvGrpSpPr/p:cNvPr/@name",
        namespaces=NS,
    )
    return found[0] if found else ""


def shape_box(element, transform: tuple[float, float, float, float]):
    xfrm = element.find("./p:spPr/a:xfrm", NS)
    if xfrm is None:
        return None
    off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
    sx, sy, tx, ty = transform
    return {
        "x": (sx * int(off.get("x")) + tx) / EMU_PER_INCH,
        "y": (sy * int(off.get("y")) + ty) / EMU_PER_INCH,
        "w": sx * int(ext.get("cx")) / EMU_PER_INCH,
        "h": sy * int(ext.get("cy")) / EMU_PER_INCH,
    }


def analyze(path: Path, legend_labels: set[str]) -> dict:
    with ZipFile(path) as package:
        root = etree.fromstring(package.read("ppt/slides/slide1.xml"))
    anchors, titles, legend_shapes, texts = {}, {}, [], []
    font_sizes = Counter()
    top_level_groups = []

    def walk(parent, transform=(1.0, 1.0, 0.0, 0.0), top_level=False):
        for element in parent:
            local = element.tag.rsplit("}", 1)[-1]
            if local == "grpSp":
                name = shape_name(element)
                if top_level and name.startswith("Editable SVG panel "):
                    top_level_groups.append(name)
                walk(element, compose(transform, group_transform(element)))
                continue
            if local != "sp":
                continue
            name = shape_name(element)
            box = shape_box(element, transform)
            text = "".join(element.xpath(".//a:t/text()", namespaces=NS)).strip()
            if text:
                texts.append(text)
            for size in element.xpath(".//a:rPr/@sz | .//a:defRPr/@sz", namespaces=NS):
                font_sizes[round(int(size) / 100, 2)] += 1
            if name.startswith("Panel layout anchor: ") and box:
                anchors[name.removeprefix("Panel layout anchor: ")] = box
            if text and re.match(r"^[A-Za-z][.)]\s*\S", text) and box:
                titles[text] = box
            if name.startswith("Shared legend marker") or name.startswith("Shared legend label"):
                legend_shapes.append({"name": name, "box": box})

    tree = root.find("./p:cSld/p:spTree", NS)
    walk(tree, top_level=True)
    return {
        "path": str(path),
        "panel_anchors": anchors,
        "panel_titles": titles,
        "editable_top_level_groups": len(top_level_groups),
        "shared_legend_shapes": legend_shapes,
        "font_sizes_pt": dict(sorted(font_sizes.items())),
        "shared_category_text_counts": dict(
            Counter(text for text in texts if text in legend_labels)
        ),
    }


def rounded_delta(candidate: dict, baseline: dict) -> dict:
    result = {}
    for key in ("x", "y", "w", "h"):
        result[key] = round(candidate[key] - baseline[key], 3)
    return result


def compare(candidate: dict, baseline: dict) -> dict:
    anchor_deltas = {}
    for name, box in candidate["panel_anchors"].items():
        if name in baseline["panel_anchors"]:
            anchor_deltas[name] = rounded_delta(box, baseline["panel_anchors"][name])
    title_deltas = {}
    for name, box in candidate["panel_titles"].items():
        if name in baseline["panel_titles"]:
            title_deltas[name] = rounded_delta(box, baseline["panel_titles"][name])
    return {
        "panel_anchor_deltas_in": anchor_deltas,
        "panel_title_deltas_in": title_deltas,
        "editable_top_level_group_delta": (
            candidate["editable_top_level_groups"] - baseline["editable_top_level_groups"]
        ),
        "shared_category_text_count_delta": {
            label: candidate["shared_category_text_counts"].get(label, 0)
            - baseline["shared_category_text_counts"].get(label, 0)
            for label in sorted(
                set(candidate["shared_category_text_counts"])
                | set(baseline["shared_category_text_counts"])
            )
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("candidate", type=Path)
    parser.add_argument("--baseline", type=Path)
    parser.add_argument("--legend-config", type=Path)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    labels = set()
    if args.legend_config:
        config = json.loads(args.legend_config.read_text(encoding="utf-8"))
        for item in config.get("items", []):
            labels.add(str(item["label"]))
            labels.update(map(str, item.get("aliases", [])))
    candidate = analyze(args.candidate, labels)
    result = {"candidate": candidate}
    if args.baseline:
        baseline = analyze(args.baseline, labels)
        result["baseline"] = baseline
        result["comparison"] = compare(candidate, baseline)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return
    print(f"Candidate: {candidate['path']}")
    print(f"Panel groups: {candidate['editable_top_level_groups']}")
    print(f"Font sizes: {candidate['font_sizes_pt']}")
    if "comparison" in result:
        print("Panel anchor deltas (candidate - baseline, inches):")
        for name, delta in result["comparison"]["panel_anchor_deltas_in"].items():
            print(f"  {name}: {delta}")
        print("Repeated shared-category labels:")
        print(f"  candidate={candidate['shared_category_text_counts']}")
        print(f"  baseline={result['baseline']['shared_category_text_counts']}")


if __name__ == "__main__":
    main()
