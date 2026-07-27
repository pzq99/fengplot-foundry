#!/usr/bin/env python3
"""Validate large-format geometry and native editability of a generated PPTX."""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}
EMU_PER_INCH = 914400
EXPECTED = 56 * EMU_PER_INCH


def local_name(element) -> str:
    return element.tag.rsplit("}", 1)[-1]


def validate(
    path: Path,
    expected_panels: int | None = None,
    expected_common_axis_anchors: int | None = None,
    expected_common_axis_size_classes: int | None = None,
    expected_shared_legend_items: int | None = None,
    expected_category_proxies: int | None = None,
    shared_category_alias_groups: list[set[str]] | None = None,
    max_shared_category_text_count: int | None = None,
    allowed_font_sizes: list[float] | None = None,
) -> dict:
    errors: list[str] = []
    with zipfile.ZipFile(path) as archive:
        presentation = etree.fromstring(archive.read("ppt/presentation.xml"))
        slide_size = presentation.find("p:sldSz", NS)
        width = int(slide_size.get("cx"))
        height = int(slide_size.get("cy"))
        if (width, height) != (EXPECTED, EXPECTED):
            errors.append(f"Canvas is {width} x {height} EMU, expected {EXPECTED} x {EXPECTED}")
        slide_names = sorted(
            n for n in archive.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")
        )
        if len(slide_names) != 1:
            errors.append(f"Expected exactly one generated slide, found {len(slide_names)}")
        slide = etree.fromstring(archive.read(slide_names[0]))
        counts = Counter(
            local_name(e)
            for e in slide.iter()
            if local_name(e) in {"sp", "grpSp", "pic", "graphicFrame", "cxnSp"}
        )
        if counts["pic"]:
            errors.append(f"Found {counts['pic']} picture objects; expected native shapes only")
        if counts["graphicFrame"]:
            errors.append(f"Found {counts['graphicFrame']} graphic frames; expected native shapes only")
        panel_groups = 0
        aligned_panel_titles = 0
        misaligned_panel_titles: list[str] = []
        common_axis_sizes: list[tuple[int, int]] = []
        for group in slide.xpath("./p:cSld/p:spTree/p:grpSp", namespaces=NS):
            c_nv_pr = group.find("p:nvGrpSpPr/p:cNvPr", NS)
            if c_nv_pr is not None and c_nv_pr.get("name", "").startswith("Editable SVG panel"):
                panel_groups += 1
                anchor_x = None
                for shape in group.findall("p:sp", NS):
                    shape_name = shape.find("p:nvSpPr/p:cNvPr", NS)
                    if shape_name is not None and shape_name.get("name", "").startswith("Panel layout anchor"):
                        anchor_off = shape.find("p:spPr/a:xfrm/a:off", NS)
                        if anchor_off is not None:
                            anchor_x = int(anchor_off.get("x"))
                        break
                for shape in group.findall("p:sp", NS):
                    shape_name = shape.find("p:nvSpPr/p:cNvPr", NS)
                    if shape_name is None or not shape_name.get("name", "").startswith("Common axes anchor"):
                        continue
                    ext = shape.find("p:spPr/a:xfrm/a:ext", NS)
                    if ext is not None:
                        common_axis_sizes.append((int(ext.get("cx")), int(ext.get("cy"))))
                    break
                for shape in group.findall("p:sp", NS):
                    text = "".join(shape.xpath(".//a:t/text()", namespaces=NS)).strip()
                    if not re.fullmatch(
                        r"(?:\([A-Za-z]\)|[A-Za-z][.)])(?:\s*\S.*)?",
                        text,
                    ):
                        continue
                    title_off = shape.find("p:spPr/a:xfrm/a:off", NS)
                    if anchor_x is not None and title_off is not None:
                        # Panel-title objects intentionally sit 0.40 in inside
                        # the full-panel boundary so italic glyphs never touch
                        # the red delivery line. All titles must use this same
                        # inset, which is the effective left-alignment rule.
                        expected_title_x = anchor_x + int(0.40 * EMU_PER_INCH)
                        delta = abs(int(title_off.get("x")) - expected_title_x)
                        if delta <= int(0.03 * EMU_PER_INCH):
                            aligned_panel_titles += 1
                        else:
                            misaligned_panel_titles.append(text)
                    break
        if panel_groups < 1:
            errors.append("No named editable SVG panel groups were found")
        if expected_panels is not None and panel_groups != expected_panels:
            errors.append(f"Found {panel_groups} panel groups, expected {expected_panels}")
        if misaligned_panel_titles:
            errors.append("Panel headings do not share the required 0.40 in left inset: " + ", ".join(misaligned_panel_titles))
        if aligned_panel_titles != panel_groups:
            errors.append(f"Only {aligned_panel_titles} of {panel_groups} panel headings passed common-inset alignment")
        if expected_common_axis_anchors is not None and len(common_axis_sizes) != expected_common_axis_anchors:
            errors.append(
                f"Found {len(common_axis_sizes)} common axes anchors, expected {expected_common_axis_anchors}"
            )
        common_axis_size_classes: list[tuple[int, int]] = []
        if common_axis_sizes:
            tolerance = int(0.03 * EMU_PER_INCH)
            for cx, cy in common_axis_sizes:
                if not any(
                    abs(cx - class_cx) <= tolerance
                    and abs(cy - class_cy) <= tolerance
                    for class_cx, class_cy in common_axis_size_classes
                ):
                    common_axis_size_classes.append((cx, cy))
            if expected_common_axis_size_classes is not None:
                if (
                    len(common_axis_size_classes)
                    != expected_common_axis_size_classes
                ):
                    errors.append(
                        "Found "
                        f"{len(common_axis_size_classes)} common axes size classes, "
                        f"expected {expected_common_axis_size_classes}"
                    )
            elif len(common_axis_size_classes) != 1:
                errors.append("Common axes anchors do not share one width and height")
        shared_legend_markers = len(
            slide.xpath('.//p:cNvPr[starts-with(@name, "Shared legend marker")]', namespaces=NS)
        )
        shared_legend_labels = len(
            slide.xpath('.//p:cNvPr[starts-with(@name, "Shared legend label")]', namespaces=NS)
        )
        if expected_shared_legend_items is not None:
            if shared_legend_markers != expected_shared_legend_items or shared_legend_labels != expected_shared_legend_items:
                errors.append(
                    "Shared legend contains "
                    f"{shared_legend_markers} markers/{shared_legend_labels} labels; "
                    f"expected {expected_shared_legend_items} each"
                )
        category_proxies = len(
            slide.xpath('.//p:cNvPr[starts-with(@name, "Category proxy marker")]', namespaces=NS)
        )
        if expected_category_proxies is not None and category_proxies != expected_category_proxies:
            errors.append(
                f"Found {category_proxies} category proxy markers, expected {expected_category_proxies}"
            )
        slide_texts = [
            "".join(shape.xpath(".//a:t/text()", namespaces=NS)).strip()
            for shape in slide.xpath(".//p:sp", namespaces=NS)
        ]
        shared_category_text_counts: dict[str, int] = {}
        for alias_group in shared_category_alias_groups or []:
            canonical = sorted(alias_group, key=lambda value: (-len(value), value))[0]
            count = sum(text in alias_group for text in slide_texts)
            shared_category_text_counts[canonical] = count
            if max_shared_category_text_count is not None and count > max_shared_category_text_count:
                errors.append(
                    f"Shared category {canonical!r} appears {count} times; "
                    f"maximum is {max_shared_category_text_count}"
                )
        if counts["sp"] < max(60, panel_groups * 20):
            errors.append(f"Found only {counts['sp']} shapes; conversion may have rasterized or dropped geometry")
        font_sizes = Counter(
            int(e.get("sz"))
            for e in slide.xpath(".//a:rPr[@sz] | .//a:endParaRPr[@sz]", namespaces=NS)
        )
        if allowed_font_sizes:
            allowed_hundredths = {round(value * 100) for value in allowed_font_sizes}
            unexpected = {
                size: count for size, count in font_sizes.items() if size not in allowed_hundredths
            }
            if unexpected:
                errors.append(
                    "Unexpected font sizes outside the allowed levels: "
                    + ", ".join(f"{size / 100:g} pt ({count})" for size, count in sorted(unexpected.items()))
                )
        elif font_sizes[8800] < max(9, panel_groups):
            errors.append(f"Only {font_sizes[8800]} text runs use the required 88 pt body size")
        top_level = slide.xpath("./p:cSld/p:spTree/*[self::p:sp or self::p:grpSp or self::p:cxnSp]", namespaces=NS)
        outside = []
        for element in top_level:
            if local_name(element) == "grpSp":
                xfrm = element.find("p:grpSpPr/a:xfrm", NS)
            else:
                xfrm = element.find("p:spPr/a:xfrm", NS)
            if xfrm is None:
                continue
            off = xfrm.find("a:off", NS)
            ext = xfrm.find("a:ext", NS)
            if off is None or ext is None:
                continue
            x, y = int(off.get("x")), int(off.get("y"))
            cx, cy = int(ext.get("cx")), int(ext.get("cy"))
            if x < 0 or y < 0 or x + cx > EXPECTED or y + cy > EXPECTED:
                c_nv_pr = element.find(".//p:cNvPr", NS)
                outside.append(c_nv_pr.get("name") if c_nv_pr is not None else local_name(element))
        if outside:
            errors.append("Top-level objects outside the 56 x 56 in canvas: " + ", ".join(outside))
        media = [
            n
            for n in archive.namelist()
            if n.startswith("ppt/media/") and not n.endswith("/")
        ]
        if media:
            errors.append("Embedded media present: " + ", ".join(media))
    report = {
        "file": str(path),
        "canvas_inches": [width / EMU_PER_INCH, height / EMU_PER_INCH],
        "slides": len(slide_names),
        "objects": dict(counts),
        "panel_groups": panel_groups,
        "aligned_panel_titles": aligned_panel_titles,
        "common_axis_anchors": len(common_axis_sizes),
        "common_axis_size_classes": len(common_axis_size_classes),
        "shared_legend_items": min(shared_legend_markers, shared_legend_labels),
        "category_proxy_markers": category_proxies,
        "shared_category_text_counts": shared_category_text_counts,
        "font_sizes_hundredths_pt": dict(font_sizes),
        "font_levels_pt": sorted(size / 100 for size in font_sizes),
        "embedded_media": media,
        "errors": errors,
        "valid": not errors,
    }
    return report


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--json", action="store_true")
    parser.add_argument("--expected-panels", type=int)
    parser.add_argument("--expected-common-axis-anchors", type=int)
    parser.add_argument("--expected-common-axis-size-classes", type=int)
    parser.add_argument("--expected-shared-legend-items", type=int)
    parser.add_argument("--expected-category-proxies", type=int)
    parser.add_argument(
        "--legend-config",
        type=Path,
        help="Legend labels and aliases used to validate repeated shared-category text",
    )
    parser.add_argument(
        "--max-shared-category-text-count",
        type=int,
        help="Maximum total visible occurrences of each legend label and its aliases",
    )
    parser.add_argument(
        "--allowed-font-sizes",
        type=float,
        nargs="+",
        help="Require every explicit text run to use one of these point sizes",
    )
    args = parser.parse_args()
    alias_groups: list[set[str]] = []
    if args.legend_config:
        config = json.loads(args.legend_config.read_text(encoding="utf-8"))
        for item in config.get("items", []):
            alias_groups.append(
                {str(item["label"]), *map(str, item.get("aliases", []))}
            )
    if args.max_shared_category_text_count is not None and not alias_groups:
        parser.error("--max-shared-category-text-count requires --legend-config")
    report = validate(
        args.pptx,
        expected_panels=args.expected_panels,
        expected_common_axis_anchors=args.expected_common_axis_anchors,
        expected_common_axis_size_classes=args.expected_common_axis_size_classes,
        expected_shared_legend_items=args.expected_shared_legend_items,
        expected_category_proxies=args.expected_category_proxies,
        shared_category_alias_groups=alias_groups,
        max_shared_category_text_count=args.max_shared_category_text_count,
        allowed_font_sizes=args.allowed_font_sizes,
    )
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"Canvas: {report['canvas_inches'][0]} x {report['canvas_inches'][1]} in")
        print(f"Objects: {report['objects']}")
        print(f"Editable panel groups: {report['panel_groups']}")
        print(f"Left-aligned panel titles: {report['aligned_panel_titles']}")
        print(f"Common axes anchors: {report['common_axis_anchors']}")
        print(
            "Common axes size classes: "
            f"{report['common_axis_size_classes']}"
        )
        print(f"Shared legend items: {report['shared_legend_items']}")
        print(f"Category proxy markers: {report['category_proxy_markers']}")
        if report["shared_category_text_counts"]:
            print(f"Shared category text counts: {report['shared_category_text_counts']}")
        print(f"Font levels: {report['font_levels_pt']}")
        print(f"Embedded media: {report['embedded_media']}")
        if report["errors"]:
            for error in report["errors"]:
                print(f"ERROR: {error}")
        else:
            print("VALID: native editable large-format PPTX")
    raise SystemExit(0 if report["valid"] else 1)


if __name__ == "__main__":
    main()
