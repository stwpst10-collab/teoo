"""Static report charts (validated palette slots 1-2; direct value labels for relief).

Numbers and labels come from chart_data.py, which build_html.py also uses for inline SVG.
"""
import sys
from pathlib import Path
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from chart_data import CHARTS  # noqa: E402

OUT = Path(sys.argv[1]); OUT.mkdir(parents=True, exist_ok=True)
plt.rcParams.update({'font.family': 'Malgun Gothic', 'axes.unicode_minus': False, 'font.size': 9})
BLUE, ORANGE = '#2a78d6', '#eb6834'
INK, MUTED, GRID, SURF = '#0b0b0b', '#5f5f5c', '#e4e4e0', '#fcfcfb'


def frame(ax):
    ax.set_facecolor(SURF)
    for s in ('top', 'right', 'left'):
        ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_color(GRID)
    ax.tick_params(colors=MUTED, length=0)


def hbar(name, labels, values, unit, xmax, note=None):
    fig, ax = plt.subplots(figsize=(6.2, 0.42 * len(labels) + 0.7), dpi=220)
    fig.patch.set_facecolor(SURF); frame(ax)
    y = range(len(labels))[::-1]
    ax.barh(list(y), values, height=0.52, color=BLUE)
    for yi, v in zip(y, values):
        ax.text(v + xmax * 0.012, yi, f'{v:g}{unit}', va='center', color=INK, fontsize=9)
    ax.set_yticks(list(y), labels, color=INK)
    ax.set_xlim(0, xmax); ax.xaxis.set_visible(False)
    ax.grid(False)
    if note:
        ax.set_title(note, loc='left', fontsize=8.5, color=MUTED, pad=6)
    fig.tight_layout(); fig.savefig(OUT / name, facecolor=SURF); plt.close(fig)


def vbar(name, labels, values, unit, ymax, note=None, decimals=1):
    fig, ax = plt.subplots(figsize=(6.2, 2.7), dpi=220)
    fig.patch.set_facecolor(SURF); frame(ax)
    x = range(len(labels))
    ax.bar(list(x), values, width=0.5, color=BLUE)
    for xi, v in zip(x, values):
        ax.text(xi, v + ymax * 0.02, f'{v:,.{decimals}f}{unit}' if v % 1 else f'{v:,.0f}{unit}', ha='center', color=INK, fontsize=9)
    ax.set_xticks(list(x), labels, color=INK)
    ax.set_ylim(0, ymax); ax.yaxis.grid(True, color=GRID, lw=0.6); ax.set_axisbelow(True)
    ax.tick_params(axis='y', labelsize=8)
    if note:
        ax.set_title(note, loc='left', fontsize=8.5, color=MUTED, pad=6)
    fig.tight_layout(); fig.savefig(OUT / name, facecolor=SURF); plt.close(fig)


def plan(name, d):
    """3-year plan: revenue and pre-tax result on one axis (c4)."""
    fig, ax = plt.subplots(figsize=(6.2, 2.9), dpi=220)
    fig.patch.set_facecolor(SURF); frame(ax)
    (n1, rev, _), (n2, pl, pl_text) = d['series']
    w = 0.32; xs = range(len(d['labels']))
    ax.bar([x - w / 2 - 0.01 for x in xs], rev, w, color=BLUE, label=n1)
    ax.bar([x + w / 2 + 0.01 for x in xs], pl, w, color=ORANGE, label=n2)
    for x, v in zip(xs, rev):
        ax.text(x - w / 2, v + 0.15, f'{v:.2f}', ha='center', color=INK, fontsize=8.5)
    for x, v, t in zip(xs, pl, pl_text):
        ax.text(x + w / 2, v + (0.15 if v >= 0 else -0.45), t, ha='center', color=INK, fontsize=8.5)
    ax.axhline(0, color=MUTED, lw=0.8)
    ax.set_xticks(list(xs), d['labels'], color=INK); ax.set_ylim(d['vmin'], d['vmax'])
    ax.yaxis.grid(True, color=GRID, lw=0.6); ax.set_axisbelow(True); ax.tick_params(axis='y', labelsize=8)
    ax.legend(frameon=False, loc='upper left', fontsize=8.5)
    ax.set_title(d['note'], loc='left', fontsize=8.5, color=MUTED, pad=6)
    fig.tight_layout(); fig.savefig(OUT / name, facecolor=SURF); plt.close(fig)


def pair(name, d):
    """Two companies side by side per year (c6)."""
    fig, ax = plt.subplots(figsize=(6.2, 2.8), dpi=220)
    fig.patch.set_facecolor(SURF); frame(ax)
    (n1, a, _), (n2, b, _) = d['series']
    w = 0.36; xs = range(len(d['labels']))
    ax.bar([x - w / 2 - 0.01 for x in xs], a, w, color=BLUE, label=n1)
    ax.bar([x + w / 2 + 0.01 for x in xs], b, w, color=ORANGE, label=n2)
    for x, v in zip(xs, a):
        ax.text(x - w / 2, v + 120, f'{v:,}', ha='center', color=INK, fontsize=7.5)
    for x, v in zip(xs, b):
        ax.text(x + w / 2, v + 120, f'{v:,}', ha='center', color=INK, fontsize=7.5)
    ax.set_xticks(list(xs), d['labels'], color=INK); ax.set_ylim(d['vmin'], d['vmax'])
    ax.yaxis.grid(True, color=GRID, lw=0.6); ax.set_axisbelow(True); ax.tick_params(axis='y', labelsize=8)
    ax.legend(frameon=False, loc='upper right', fontsize=8.5, ncol=2)
    ax.set_title(d['note'], loc='left', fontsize=8.5, color=MUTED, pad=6)
    fig.tight_layout(); fig.savefig(OUT / name, facecolor=SURF); plt.close(fig)


for name, d in CHARTS.items():
    if d['kind'] == 'hbar':
        hbar(name, d['labels'], d['values'], d['unit'], d['vmax'], d['note'])
    elif d['kind'] == 'vbar':
        vbar(name, d['labels'], d['values'], d['unit'], d['vmax'], d['note'], d.get('decimals', 1))
    elif name == 'c4_plan.png':
        plan(name, d)
    else:
        pair(name, d)
print('charts ok', len(CHARTS))
