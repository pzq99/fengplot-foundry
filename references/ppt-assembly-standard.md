# Standard for final editable PowerPoint assembly

Use this standard to choose the layout, convert panels, and accept the final figure.

## Delivery canvas

- Slide: `56 × 56 in`.
- Primary horizontal region: `x=2–54 in`, width `52 in`.
- Title band: `x=2 in`, `y=0`, `w=52 in`, `h≈2.98 in`, fill `#785F8E`.
- Slide title: Microsoft YaHei, `120 pt`, white.
- Reference boundary: red `10 pt` box at `x=2`, `y=0`, `w=52`, `h=52`.

Fill the horizontal region. Do not force the assembled panels to fill the red square vertically. A compact figure with deliberate white space below is preferable to stretched or oversized panels.

## Layout selection

Choose row counts from semantic grouping and visible information density, not panel count alone.

- Use `1/4/4` or `1/3/4` when a wide overview panel leads the figure.
- Use `2/3/3`, `3/3/2`, or another balanced arrangement for plot-only figures.
- Keep comparable coordinate-axis plots in the same row when possible.
- Reorder only when the user allows it and the scientific narrative remains accurate.
- Place dense panels beside sparse panels only when their visual masses remain balanced.
- Make the first and last panel of every populated row touch `x=2` and `x=54`.
- Reduce gaps before enlarging text or distorting plots.

For a compact `1/4/4` figure, start near `0.01 in` horizontal and `0.25 in` vertical panel-box gaps. Judge visible whitespace inside the SVGs separately from outer box gaps.

Use `--fit-rows-to-width` for a sparse overview row that must touch both horizontal boundaries without changing aspect ratio. Use `--pack-rows` when the composition should remain vertically compact.

## Information density

Evaluate visible ink, not outer SVG dimensions.

- Let a concise overview panel be wide and relatively shallow.
- Enlarge lower-row plots until axes and important marks dominate each cell.
- Align panel headings, axes tops, axes bottoms, and role-equivalent text.
- Avoid one dense panel carrying the row while its neighbors contain large unused margins.
- Prefer shortening text, widening text boxes, or moving exact detail to notes over shrinking the whole figure.

## Typography

Use a fixed role system derived from the final manually refined reference:

| PowerPoint role | Size |
|---|---:|
| Slide title | 120 pt |
| Panel heading | 76 pt |
| Overview section heading | 66 pt |
| Axis label | 60 pt |
| Body, tick, annotation, legend | 48 pt |
| Micro detail | 40 pt only when essential |

Do not introduce adjacent sizes such as 44, 52, 54, 58, or 64 pt. Use no more than the roles required by the figure.

For standardized source SVGs, use:

```bash
--font-size 76 \
--preserve-source-font-ratios \
--preserved-ratio-role-levels 66 60 48 40
```

Keep panel headings bold italic and axis labels italic. Keep body text upright unless scientific notation requires otherwise.

## Axes and aspect ratios

Read every SVG viewBox and preserve its aspect ratio.

- Normalize axes only within the same `alignment_group`.
- Give comparable plots equal axes height and, when possible, equal axes width.
- Align spines rather than only panel bounds.
- Exempt overview panels, heatmaps, and dendrogram composites.
- Use one physical axis/tick stroke width across differently scaled panels; `5 pt` is the validated starting point.
- Use a `2.5 pt` minimum for colored data strokes.

Do not use nonuniform scaling to make a plot fill its cell.

## Overview panel treatment

Place a wide overview panel across the first row when it provides the figure's narrative.

- Align its outer left and right edges to the delivery boundaries.
- Keep its height driven by content, not by the square canvas.
- Use three or fewer major stages.
- Keep frames visibly heavier than internal diagram strokes.
- Keep the section headings large enough to scan with the lower-row panel headings.
- Keep body wording brief; retain full method detail in slide notes.
- Preserve consistent card widths, internal padding, and baseline alignment.

## Legends and label collision

Show each shared category name once.

- Build each marker–label pair as one logical unit.
- Give text boxes enough width to remain on one line.
- Align legend items to a common row or column before grouping them.
- Order items to match plots and heatmap columns.
- For a multiline legend, center a short final item instead of forcing it into a collision.
- After conversion, inspect every legend at 100% zoom. If items overlap, expand the band or reduce gaps elsewhere; do not shrink only one label.

The manually refined reference regrouped the `d`-panel legend as five independent marker–label units. Treat that as the reusable rule; do not copy the incidental ungrouping of unrelated panels.

## Editability and grouping

Convert SVG paths, rectangles, circles, markers, lines, and text into native PowerPoint objects. Preserve SVG order as z-order.

- Reject `<image>` in editable mode.
- Use zero `p:pic` and zero embedded media.
- Group a complete panel once for movement.
- Keep marker–label legend items grouped inside the panel where practical.
- Ensure every child lies inside the panel group's coordinate bounds.
- Ungroup and regroup during QA to verify that text and geometry do not jump or clip.
- If grouping truncates labels, fix the child/group bounds in the converter. Do not accept ungrouping as the only delivery fix.

## Acceptance workflow

1. Validate the source SVGs.
2. Convert into a new PPTX filename.
3. Run structural validation with the declared font levels.
4. Render with Microsoft PowerPoint to PDF and PNG.
5. Inspect full-resolution PNG for:
   - all text fully visible;
   - no overlaps;
   - left/right boundary contact;
   - balanced row density;
   - common axes alignment;
   - correct legend order and grouping;
   - no excessive vertical stretching;
   - no objects outside the 56-inch canvas.
6. If a manual final exists, compare it to a regenerated baseline and promote only coherent repeated differences.

Quick Look is not authoritative. Deliver the editable PPTX plus a matching PowerPoint-rendered PDF/PNG preview when requested.
