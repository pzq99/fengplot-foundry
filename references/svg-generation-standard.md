# Standard for code-generated SVG panels

Apply this standard to every Python/Matplotlib panel before PowerPoint assembly.

## Deliverables

Generate and run one deterministic script, normally `generate_panels.py`. Deliver:

- `panel-01.svg` through `panel-NN.svg` in display order;
- `manifest.json`;
- `validation-report.txt`;
- the runnable source script and any non-confidential input data or data paths.

Create each panel as a separate figure. Never export a preassembled grid.

## Required Matplotlib setup

Use a noninteractive backend before importing `pyplot`. Keep text native, the background transparent, and layout margins explicit.

```python
import matplotlib as mpl

mpl.use("Agg")
mpl.rcParams.update({
    "svg.fonttype": "none",
    "font.family": "Arial",
    "font.size": 24,
    "axes.titlesize": 24,
    "axes.labelsize": 30,
    "axes.linewidth": 2.5,
    "xtick.labelsize": 24,
    "ytick.labelsize": 24,
    "xtick.major.size": 8,
    "ytick.major.size": 8,
    "xtick.major.width": 2.5,
    "ytick.major.width": 2.5,
    "legend.fontsize": 24,
    "lines.linewidth": 4.0,
    "patch.linewidth": 2.2,
})
```

Save with:

```python
fig.savefig(path, format="svg", transparent=True)
```

Do not use `bbox_inches="tight"`, `tight_layout()`, `constrained_layout=True`, `rasterized=True`, `text.usetex=True`, or `imshow`.

## Typography roles

Use only these source roles unless the content has a documented exception:

| Role | Source size | Purpose |
|---|---:|---|
| Panel heading | 40 pt | `a.`, `b.`, … heading |
| Section heading | 34 pt | Overview stages, major callouts |
| Axis label | 30–32 pt | x/y labels and other primary quantitative labels |
| Body/tick/legend | 24 pt | Ticks, annotations, standard legends, diagram body |
| Micro detail | 18–20 pt | Rare secondary evidence that cannot be removed |

Do not create intermediate one-off sizes. Prefer deleting redundant words or moving exact detail to notes over using micro text.

Add one figure-level heading:

```python
def add_panel_title(fig, title: str) -> None:
    fig.text(
        0.0, 0.985, title,
        ha="left", va="top",
        fontsize=40,
        fontstyle="italic",
        fontweight="bold",
    )
```

Use consecutive `a.`, `b.`, `c.` prefixes. Do not use an axes title as the panel heading.

## Canvas, axes, and alignment

Use the same `figsize` and explicit margins for comparable plots. A reliable standard plot starting point is:

```python
fig = plt.figure(figsize=(6.5, 6.0))
fig.subplots_adjust(left=0.24, right=1.0, bottom=0.26, top=0.89)
```

Tune the values once for the figure family, then reuse them. Do not shorten one axes merely to fit a repeated legend.

Use wider or taller canvases only for true semantic exceptions such as a workflow overview, heatmap, or dendrogram composite. Record each exception in `manifest.json`. Preserve comparable axes height even when widths differ.

Keep all labels, colorbars, legends, and annotations inside the viewBox. Avoid post-export cropping.

## Overview and workflow panels

Build a top overview panel as a concise vector diagram, not a paragraph-heavy infographic.

- Use one dominant left-to-right reading direction.
- Limit the top-level structure to about three stages.
- Use aligned outer frames and aligned internal cards.
- Keep the left and right content edges intentional so the panel can span a full PowerPoint row.
- Use short noun phrases and one quantitative takeaway per stage.
- Encode methods and stages with consistent color, shape, and stroke.
- Prefer a wide, shallow canvas. A `3.0–3.4:1` aspect ratio is a useful starting range.
- Use `4–4.5 pt` main frames and `2.5–3 pt` internal strokes at the source scale when the panel will occupy the full 52-inch width.

Remove low-value decorative elements and repeated explanations. Put full methodological wording in notes or a companion text file.

## Plot-specific requirements

### Lines, bars, and scatter

- Keep line vertices and scatter markers to the minimum needed for the scientific claim.
- Use vector rectangles for bars and vector markers for scatter.
- Keep grid lines lighter than data.
- Use both color and line/marker semantics when curves must remain distinguishable.

### Heatmaps and clustered heatmaps

- Use `pcolormesh`, `pcolor`, or explicit rectangle patches.
- Use a discrete `BoundaryNorm`; keep colorbar segments under 50.
- Draw dendrograms as vector paths and reorder the matrix by leaf order before plotting.
- Keep matrices at an editable object count; aggregate or confirm before exceeding about 20 × 20 cells.

### Legends

- Keep a panel-local legend only when its semantics are unique.
- When two or more panels share a mapping, write the mapping once in `manifest.json.shared_legend` and omit repeated legend text from the SVGs.
- Preserve item order across plots, heatmap columns, and the final shared legend.

## Manifest

Use at least:

```json
{
  "seed": 20260718,
  "panel_count": 3,
  "shared_legend": {
    "position": "after-first-row",
    "items": [
      {"label": "Method A", "color": "#7B61A8", "marker": "circle"}
    ]
  },
  "panels": [
    {
      "index": 1,
      "letter": "a",
      "file": "panel-01.svg",
      "title": "a. Task overview",
      "plot_type": "workflow",
      "viewBox": [0, 0, 1680, 520],
      "aspect_ratio": 3.230769,
      "alignment_group": "overview",
      "axis_exception": true
    }
  ]
}
```

Set `alignment_group` consistently for panels whose axes must align. Mark only genuine exceptions with `axis_exception: true`.

## Mandatory QA

Run:

```bash
python3 scripts/validate_svg_panels.py /path/to/svg-panels
```

Also render or open every SVG and inspect:

- title continuity and left anchoring;
- text overlap and clipping;
- identical axes boxes within each alignment group;
- readable ticks and axis labels;
- absence of repeated legends;
- balanced visible density;
- no `<image>`, base64 bitmap, external image, or all-text-as-path output.

Regenerate from code after every correction. Do not hand-edit the SVG as the primary fix.
