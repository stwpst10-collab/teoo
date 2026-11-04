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


def expansion(name, d):
    """Stage map (no numbers): one blue ramp light→dark for stage order, gray only for the contrast box."""
    from matplotlib.patches import FancyBboxPatch
    ramp = ['#d6e6f8', '#adcdf1', '#7fb1e8', '#4f8fdc']
    fig = plt.figure(figsize=(6.6, 7.0), facecolor=SURF)
    ax = fig.add_axes([0.01, 0.01, 0.98, 0.98]); ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis('off')
    br = lambda s: s.replace('|', '\n')  # noqa: E731
    ax.text(0, 0.985, d['note'], fontsize=8, color=MUTED, va='top')
    ax.text(1, 0.985, '● 기존 근거   ○ 조사 중', fontsize=8, color=MUTED, va='top', ha='right')
    ax.text(0, 0.955, '→ 조건을 채우면 다음 단계. ' + d['gate'], fontsize=8, color=MUTED, va='top')
    top, bottom, gap = 0.915, 0.11, 0.03
    ws = [0.2, 0.2, 0.2, 0.31]
    xs = [sum(ws[:i]) + gap * i for i in range(4)]

    def box(x, y0, y1, st, color, compact):
        w = ws[3] if compact else ws[0]
        ax.add_patch(FancyBboxPatch((x, y0), w, y1 - y0, boxstyle='round,pad=0,rounding_size=0.008',
                                    fc='white', ec=GRID, lw=1))
        ax.add_patch(FancyBboxPatch((x, y1 - 0.012), w, 0.012, boxstyle='square,pad=0', fc=color, ec='none'))
        mark = '●' if st['evidence'] else '○'
        ax.text(x + 0.012, y1 - 0.02, f"{st['key']} {br(st['title'])}", fontsize=8.5, weight='bold', color=INK, va='top')
        ax.text(x + w - 0.01, y1 - 0.02, mark, fontsize=9, color=BLUE, va='top', ha='right')
        y = y1 - (0.078 if '|' in st['title'] else 0.052)
        step = (y - y0 - 0.01) / len(d['rows'])
        for lab, val in zip(d['rows'], st['cells']):
            ax.text(x + 0.012, y, lab, fontsize=6.5, color=MUTED, va='top')
            ax.text(x + 0.012, y - (0.018 if compact else 0.022), br(val), fontsize=7 if compact else 7.5,
                    color=INK, va='top', linespacing=1.25)
            y -= step
        return (x, (y0 + y1) / 2)

    mids = [box(xs[i], bottom, top, d['stages'][i], ramp[i], False)[1] for i in range(3)]
    h = (top - bottom - 2 * 0.015) / 3
    four = []
    for k in range(3):
        y1 = top - k * (h + 0.015)
        four.append(box(xs[3], y1 - h, y1, d['stages'][3 + k], ramp[3], True)[1])
    arrow = dict(arrowstyle='-|>', color=MUTED, lw=1)
    for i in range(2):
        ax.annotate('', xy=(xs[i + 1] - 0.004, mids[i + 1]), xytext=(xs[i] + ws[i] + 0.004, mids[i]), arrowprops=arrow)
    for ym in four:
        ax.annotate('', xy=(xs[3] - 0.004, ym), xytext=(xs[2] + ws[2] + 0.004, mids[2]), arrowprops=arrow)
    ax.add_patch(FancyBboxPatch((0, 0.02), 1, 0.05, boxstyle='round,pad=0,rounding_size=0.008', fc=GRID, ec='none'))
    ax.text(0.5, 0.045, d['contrast'], fontsize=8, color=MUTED, ha='center', va='center')
    fig.savefig(OUT / name, facecolor=SURF, dpi=200); plt.close(fig)


for name, d in CHARTS.items():
    if d['kind'] == 'map':
        expansion(name, d)
    elif d['kind'] == 'hbar':
        hbar(name, d['labels'], d['values'], d['unit'], d['vmax'], d['note'])
    elif d['kind'] == 'vbar':
        vbar(name, d['labels'], d['values'], d['unit'], d['vmax'], d['note'], d.get('decimals', 1))
    elif name == 'c4_plan.png':
        plan(name, d)
    else:
        pair(name, d)
print('charts ok', len(CHARTS))
