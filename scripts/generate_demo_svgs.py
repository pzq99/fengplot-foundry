#!/usr/bin/env python3
"""Generate deterministic Matplotlib SVG panels for end-to-end skill testing."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

import matplotlib as mpl

mpl.use("Agg")
mpl.rcParams.update(
    {
        "svg.fonttype": "none",
        "font.family": "Arial",
        "font.size": 24,
        "axes.titlesize": 34,
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
    }
)

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.colors import BoundaryNorm
from scipy.cluster.hierarchy import dendrogram, leaves_list, linkage


FIGSIZE = (6.5, 6.0)
COLORS = {
    "blue": "#00A6D6",
    "green": "#86C980",
    "orange": "#FFB45B",
    "red": "#F44336",
    "purple": "#9C27B0",
    "slate": "#607D8B",
    "gold": "#F2A93B",
}


def add_panel_title(fig, title: str) -> None:
    fig.text(
        0.0,
        0.985,
        title,
        ha="left",
        va="top",
        fontstyle="italic",
        fontweight="bold",
        fontsize=40,
    )


def new_axes(
    title: str,
    *,
    figsize=FIGSIZE,
    margins=(0.21, 0.98, 0.20, 0.86),
    grid_axis="y",
):
    fig, ax = plt.subplots(figsize=figsize)
    left, right, bottom, top = margins
    fig.subplots_adjust(left=left, right=right, bottom=bottom, top=top)
    add_panel_title(fig, title)
    if grid_axis:
        ax.grid(axis=grid_axis, color="#D9D9D9", linestyle="--", linewidth=2.0, alpha=0.85)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    return fig, ax


def save(fig, path: Path) -> None:
    fig.savefig(path, format="svg", transparent=True, metadata={"Creator": "Matplotlib demo generator"})
    plt.close(fig)


def panel_a(rng, path):
    fig, ax = new_axes("a. Ligand atom count")
    x = rng.gamma(4.2, 5.4, 4200) + 5
    ax.hist(x, bins=28, color=COLORS["blue"], edgecolor="#111111")
    med, mean = np.median(x), np.mean(x)
    ax.axvline(med, color="#4CAF50", linewidth=4)
    ax.axvline(mean, color=COLORS["red"], linestyle="--", linewidth=4)
    ax.text(med - 1.0, ax.get_ylim()[1] * 0.91, f"{med:.1f}", color="#4CAF50", ha="right")
    ax.text(mean + 1.0, ax.get_ylim()[1] * 0.80, f"{mean:.1f}", color=COLORS["red"], ha="left")
    ax.set_xlabel("Number of ligand atoms", fontstyle="italic")
    ax.set_ylabel("Count", fontstyle="italic")
    ax.set_xticks([20, 40, 60, 80, 100])
    save(fig, path)


def panel_b(rng, path):
    fig, ax = new_axes("b. QED distribution")
    x = np.concatenate([rng.beta(3.2, 4.8, 2700), rng.beta(7.0, 2.5, 1400)])
    ax.hist(x, bins=30, density=True, color=COLORS["green"], edgecolor="#111111")
    med, mean = np.median(x), np.mean(x)
    ax.axvline(med, color="#4CAF50", linewidth=4)
    ax.axvline(mean, color=COLORS["red"], linestyle="--", linewidth=4)
    ax.text(med - 0.035, ax.get_ylim()[1] * 0.91, f"{med:.2f}", color="#4CAF50", ha="right")
    ax.text(mean + 0.035, ax.get_ylim()[1] * 0.80, f"{mean:.2f}", color=COLORS["red"], ha="left")
    ax.set_xlabel("QED", fontstyle="italic")
    ax.set_ylabel("Density", fontstyle="italic")
    ax.set_xlim(0, 1)
    ax.set_xticks([0, 0.25, 0.5, 0.75, 1.0])
    save(fig, path)


def panel_c(rng, path):
    fig, ax = new_axes("c. Pocket residues")
    x = np.concatenate([rng.normal(28, 10, 1300), rng.normal(58, 17, 3000)])
    x = x[(x > 5) & (x < 125)]
    ax.hist(x, bins=26, color=COLORS["green"], edgecolor="#111111")
    med = np.median(x)
    ax.axvline(med, color=COLORS["red"], linestyle="--", linewidth=4)
    ax.text(med + 4, ax.get_ylim()[1] * 0.91, f"{med:.0f}", color=COLORS["red"])
    ax.set_xlabel("Pocket residue count", fontstyle="italic")
    ax.set_ylabel("Count", fontstyle="italic")
    ax.set_xlim(0, 130)
    ax.set_xticks([20, 40, 60, 80, 100, 120])
    save(fig, path)


def panel_d(rng, path):
    fig, ax = new_axes("d. Pocket volume")
    x = np.concatenate([rng.gamma(2.0, 900, 1200), rng.normal(5400, 2100, 3200)])
    x = x[(x > 300) & (x < 15000)]
    ax.hist(x, bins=28, color=COLORS["orange"], edgecolor="#111111")
    med = np.median(x)
    ax.axvline(med, color=COLORS["red"], linestyle="--", linewidth=4)
    ax.text(med + 500, ax.get_ylim()[1] * 0.91, f"{med:.0f}", color=COLORS["red"])
    ax.set_xlabel("Pocket volume", fontstyle="italic")
    ax.set_ylabel("Count", fontstyle="italic")
    ax.set_xticks([2000, 6000, 10000, 14000])
    save(fig, path)


def panel_e(rng, path):
    fig, ax = new_axes(
        "e. Training loss",
        figsize=(8.0, 5.4),
        margins=(0.17, 0.98, 0.22, 0.84),
    )
    t = np.arange(1, 51)
    for label, color, scale in [
        ("Gaussian", "#42A6F5", 1.22),
        ("Sobol", "#4CAF50", 1.05),
        ("Shell", "#FF9800", 0.90),
    ]:
        y = scale * (11 / np.sqrt(t) + 0.35) + rng.normal(0, 0.05, len(t))
        ax.plot(t, y, color=color, label=label)
    ax.set_xlabel("Epoch", fontstyle="italic")
    ax.set_ylabel("Train loss", fontstyle="italic")
    ax.set_ylim(0, 15)
    ax.set_yticks([0, 5, 10, 15])
    ax.legend(frameon=False, loc="upper right")
    save(fig, path)


def panel_f(rng, path):
    fig, ax = new_axes(
        "f. RMSD relationship",
        figsize=(6.0, 6.6),
        margins=(0.23, 0.98, 0.18, 0.87),
    )
    x = rng.normal(0, 1.0, 110)
    y = 0.62 * x + rng.normal(0, 0.62, len(x))
    ax.scatter(x, y, s=100, color="#78909C", edgecolor="#263238", alpha=0.85)
    coef = np.polyfit(x, y, 1)
    xx = np.linspace(-2.8, 2.8, 80)
    ax.plot(xx, np.polyval(coef, xx), color=COLORS["red"])
    ax.set_xlabel("Predicted RMSD", fontstyle="italic")
    ax.set_ylabel("Observed RMSD", fontstyle="italic")
    ax.set_xticks([-2, 0, 2])
    ax.set_yticks([-2, 0, 2])
    save(fig, path)


def panel_g(rng, path):
    fig, ax = new_axes(
        "g. Correlation heatmap",
        figsize=(7.6, 6.0),
        margins=(0.16, 0.78, 0.38, 0.84),
        grid_axis=None,
    )
    latent = rng.normal(size=(500, 3))
    weights = rng.normal(size=(3, 6))
    data = np.einsum("ij,jk->ik", latent, weights) + rng.normal(
        0, 0.55, size=(500, 6)
    )
    matrix = np.corrcoef(data, rowvar=False)
    boundaries = np.linspace(-1, 1, 8)
    heatmap_cmap = plt.get_cmap("RdBu_r", len(boundaries) - 1)
    heatmap_norm = BoundaryNorm(boundaries, heatmap_cmap.N)
    mesh = ax.pcolormesh(
        np.arange(7),
        np.arange(7),
        matrix,
        cmap=heatmap_cmap,
        norm=heatmap_norm,
        shading="flat",
        edgecolors="#FFFFFF",
        linewidth=1.2,
    )
    labels = [f"V{i}" for i in range(1, 7)]
    ax.set_xticks(np.arange(6) + 0.5, labels=labels, rotation=35, ha="right")
    ax.set_yticks(np.arange(6) + 0.5, labels=labels)
    ax.set_xlabel("Variables", fontstyle="italic", labelpad=24)
    ax.set_ylabel("Variables", fontstyle="italic")
    colorbar = fig.colorbar(
        mesh,
        ax=ax,
        fraction=0.055,
        pad=0.04,
        boundaries=boundaries,
        ticks=[-1, 0, 1],
    )
    colorbar.set_label("Correlation", fontstyle="italic")
    save(fig, path)


def panel_h(rng, path):
    fig = plt.figure(figsize=(8.4, 6.3))
    add_panel_title(fig, "h. Hierarchical clustering")
    grid = fig.add_gridspec(
        2,
        2,
        width_ratios=(1.05, 4.8),
        height_ratios=(1.05, 4.8),
        left=0.08,
        right=0.98,
        bottom=0.36,
        top=0.84,
        wspace=0.03,
        hspace=0.03,
    )
    ax_corner = fig.add_subplot(grid[0, 0])
    ax_top = fig.add_subplot(grid[0, 1])
    ax_left = fig.add_subplot(grid[1, 0])
    ax_heat = fig.add_subplot(grid[1, 1])
    ax_corner.axis("off")

    samples = rng.normal(size=(8, 8))
    samples[:3, :3] += 2.2
    samples[3:6, 3:6] -= 1.8
    samples[6:, 6:] += 1.5
    row_tree = linkage(samples, method="ward")
    column_tree = linkage(samples.T, method="ward")
    row_order = leaves_list(row_tree)
    column_order = leaves_list(column_tree)
    dendrogram(
        column_tree,
        ax=ax_top,
        orientation="top",
        no_labels=True,
        color_threshold=0,
        above_threshold_color="#607D8B",
    )
    dendrogram(
        row_tree,
        ax=ax_left,
        orientation="left",
        no_labels=True,
        color_threshold=0,
        above_threshold_color="#607D8B",
    )
    ordered = samples[np.ix_(row_order, column_order)]
    ax_heat.pcolormesh(
        np.arange(9),
        np.arange(9),
        ordered,
        cmap="PuOr_r",
        shading="flat",
        edgecolors="#FFFFFF",
        linewidth=0.8,
    )
    ax_heat.set_xticks(np.arange(8) + 0.5, labels=[f"F{i + 1}" for i in column_order], rotation=45, ha="right")
    ax_heat.set_yticks(np.arange(8) + 0.5, labels=[f"S{i + 1}" for i in row_order])
    ax_heat.set_xlabel("Features", fontstyle="italic", labelpad=24)
    for axis in (ax_top, ax_left):
        axis.set_xticks([])
        axis.set_yticks([])
        for spine in axis.spines.values():
            spine.set_visible(False)
    save(fig, path)


def panel_i(rng, path):
    fig, ax = new_axes("i. Calibration curves")
    t = np.linspace(0, 1, 9)
    for offset, color, label in [
        (0.0, "#42A6F5", "Baseline"),
        (0.14, "#4CAF50", "Balanced"),
        (-0.12, "#FF9800", "Weighted"),
    ]:
        y = (t - 0.52 - offset) ** 2 * 2.7 + 0.2 + rng.normal(0, 0.035, len(t))
        ax.plot(t, y, marker="o", markersize=10, color=color, label=label)
    ax.set_xlabel("Confidence", fontstyle="italic")
    ax.set_ylabel("Calibration loss", fontstyle="italic")
    ax.set_xlim(-0.05, 1.08)
    ax.set_ylim(0, 2.3)
    ax.set_xticks([0, 0.25, 0.5, 0.75, 1.0])
    ax.legend(frameon=False, loc="upper right")
    save(fig, path)


PANELS = [panel_a, panel_b, panel_c, panel_d, panel_e, panel_f, panel_g, panel_h, panel_i]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--seed", type=int, default=20260718)
    parser.add_argument("--count", type=int, default=9)
    args = parser.parse_args()
    if not 2 <= args.count <= len(PANELS):
        parser.error(f"--count must be between 2 and {len(PANELS)}")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    for stale in args.output_dir.glob("panel-*.svg"):
        stale.unlink()
    rng = np.random.default_rng(args.seed)
    manifest = []
    for index, func in enumerate(PANELS[: args.count], start=1):
        output = args.output_dir / f"panel-{index:02d}.svg"
        func(rng, output)
        manifest.append({"index": index, "file": output.name, "generator": func.__name__})
    (args.output_dir / "manifest.json").write_text(
        json.dumps(
            {
                "seed": args.seed,
                "panel_count": len(manifest),
                "panels": manifest,
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"Generated {len(manifest)} SVG panels in {args.output_dir}")


if __name__ == "__main__":
    main()
