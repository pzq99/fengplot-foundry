---
name: fengplot-foundry
description: "Apply the Zhu Feng research group's figure-making practice as a two-stage workflow: generate standardized, pure-vector scientific panels with Python/Matplotlib, then forge 2–12 panels into a polished, large-format, fully editable PowerPoint figure. Use when Codex must create or normalize SVG plots, build diagrammatic overview panels, choose a dense multi-row layout, align comparable axes, preserve native text and marks, consolidate legends, reproduce a manually refined figure, or validate an SVG-to-editable-PPTX workflow on the bundled 56 × 56 inch template."
---

# FengPlot Foundry

Complete the workflow in two parts. Do not insert SVGs as pictures and do not rasterize the final figure.

## Scope and permitted use

- This skill is designed and tested for OpenAI Codex. Other models, agent frameworks, and runtime environments have not been tested.
- This skill is for personal use only.
- Select the smallest workflow that satisfies the task: run Part I only to generate or normalize Matplotlib SVG panels, Part II only to assemble existing SVG panels into an editable PowerPoint, or both parts for an end-to-end build.

## Part I — Generate standardized SVG panels

Read [references/svg-generation-standard.md](references/svg-generation-standard.md) before writing or revising panel-generation code.

1. Create one independent SVG per panel and name files `panel-01.svg`, `panel-02.svg`, and so on in display order.
2. Generate SVGs from a deterministic Python script. Keep all text as `<text>`, all marks as vectors, and all canvases transparent.
3. Use the five source typography roles defined in the reference: panel heading, section heading, axis label, body/tick, and optional micro detail. Do not invent one-off sizes.
4. Give comparable plots identical figure size, explicit subplot margins, and identical axes geometry. Declare genuine aspect-ratio exceptions in `manifest.json`.
5. Build overview or workflow panels as concise vector diagrams. Remove redundant prose before shrinking text.
6. Validate the directory:

```bash
python3 scripts/validate_svg_panels.py /path/to/svg-panels \
  --manifest /path/to/svg-panels/manifest.json
```

Stop and regenerate if validation finds raster images, invalid viewBoxes, missing native text, discontinuous filenames, or incorrect panel headings.

For a runnable example, inspect [demo/svg-panels](demo/svg-panels) and regenerate it with `scripts/generate_demo_svgs.py`.

## Part II — Assemble the editable PowerPoint

Read [references/ppt-assembly-standard.md](references/ppt-assembly-standard.md) before choosing a layout or converting the SVGs.

Use [assets/standard-template.pptx](assets/standard-template.pptx) unless the user supplies a verified equivalent 56 × 56 inch template.

1. Inventory panel aspect ratios, axes boxes, legends, visible density, and semantic relationships.
2. Choose rows by narrative and information density. Make every populated row reach the left and right delivery boundaries; do not force the figure to fill the full square vertically.
3. Preserve aspect ratios. Normalize axes only within comparable plot classes.
4. Convert paths, lines, markers, fills, and text to native PowerPoint objects. Preserve source element order as z-order.
5. Use a small role-based type system. For standardized SVGs, map source roles to the manually validated PowerPoint levels:

```bash
--font-size 76 \
--preserve-source-font-ratios \
--preserved-ratio-role-levels 66 60 48 40
```

This yields `76 pt` panel headings, `66 pt` section headings, `60 pt` axis labels, `48 pt` body/ticks/legends, and `40 pt` micro detail. The slide title remains `120 pt`. Omit the micro tier when it is unnecessary.

6. Build repeated legends once. Treat each marker–label pair as one logical unit, align the units, and group them only after collision checks.
7. Group each panel for convenient movement only when the group bounds contain every child and PowerPoint renders it without clipping. Fix group bounds rather than accepting truncated labels.
8. Save to a new filename when replacing a cached or manually edited version.

Typical dense `1/4/4` composition:

```bash
python3 scripts/svg_to_editable_pptx.py \
  --template assets/standard-template.pptx \
  --svg-dir /path/to/svg-panels \
  --output /path/to/output-new.pptx \
  --title "Fig 3" \
  --font Calibri \
  --font-size 76 \
  --preserve-source-font-ratios \
  --preserved-ratio-role-levels 66 60 48 40 \
  --max-columns 4 \
  --row-counts 1 4 4 \
  --fit-rows-to-width 1 \
  --pack-rows \
  --min-gap-x 0.01 \
  --min-gap-y 0.25 \
  --normalize-common-axes \
  --axis-stroke-width 5 \
  --min-stroke-width 2.5
```

Treat the numeric gaps and axes box as starting points, not universal constants. Inspect the rendered slide and adjust from evidence.

The packaged end-to-end output is [demo/fengplot-foundry-demo-editable.pptx](demo/fengplot-foundry-demo-editable.pptx).

## Validation and learning from a manual final

Run structural validation:

```bash
python3 scripts/validate_editable_pptx.py /path/to/output-new.pptx \
  --expected-panels 9 \
  --allowed-font-sizes 40 48 60 66 76 120
```

Render with Microsoft PowerPoint on macOS:

```bash
python3 scripts/render_with_powerpoint.py /path/to/output-new.pptx \
  --pdf /path/to/output-new.pdf \
  --png /path/to/output-new.png
```

Inspect the PNG at full resolution for clipping, overlaps, boundary contact, axes alignment, legend spacing, and density balance. Quick Look is not authoritative.

When a manually refined PPTX exists, regenerate the previous automatic baseline and compare:

```bash
python3 scripts/compare_manual_pptx_layout.py /path/to/manual-final.pptx \
  --baseline /path/to/regenerated-baseline.pptx
```

Promote coherent changes into rules: repeated font-role changes, row-level movement, consistent boundary contact, legend regrouping, or systematic label expansion. Do not promote accidental ungrouping, isolated nudges, split text runs, or font substitution.

## Operating constraints

- Write cloud-volume output through a local temporary PPTX and copy only the completed package.
- Reject source `<image>` elements in fully editable mode.
- Use vector heatmap cells and dendrogram paths; never use `imshow`.
- Keep every object within the 56 × 56 inch canvas.
- Keep exact values omitted for visual clarity in notes or companion JSON/CSV.
- Do not deliver until structural validation and PowerPoint visual QA both pass.
