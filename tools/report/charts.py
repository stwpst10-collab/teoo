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
    ax.text(1, 0.985, '● 규모 근거 있음   ○ 규모 근거 없음', fontsize=8, color=MUTED, va='top', ha='right')
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


ROAD_FILL = {'go': ('#d6e6f8', BLUE, INK), 'try': ('white', BLUE, INK),
             'turn': ('#fbe1d6', ORANGE, INK), 'hold': ('#efefeb', '#efefeb', MUTED)}


def roadmap(name, d):
    """Scenario × stage grid: rows are scenarios, columns the same stages; cell state says how each scenario passes."""
    from matplotlib.patches import FancyBboxPatch
    br = lambda s: s.replace('|', '\n')  # noqa: E731
    n = len(d['stages'])
    fig = plt.figure(figsize=(7.0, 5.6), facecolor=SURF)
    ax = fig.add_axes([0.005, 0.005, 0.99, 0.99]); ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis('off')
    lw_, rw, gap = 0.1, 0.12, 0.008
    cw = (1 - lw_ - rw - gap * (n + 1)) / n
    cx = [lw_ + gap + i * (cw + gap) for i in range(n)]
    rx = cx[-1] + cw + gap

    def cell(x, y, w, h, state, text, fs=7, bold=False, dashed=False):
        fc, ec, tc = ROAD_FILL[state]
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0,rounding_size=0.006', fc=fc, ec=ec,
                                    lw=1.1 if state in ('go', 'try', 'turn') else 0, ls='--' if dashed or state == 'try' else '-'))
        ax.text(x + w / 2, y + h / 2, br(text), fontsize=fs, color=tc, ha='center', va='center',
                linespacing=1.3, weight='bold' if bold else 'normal')

    # legend + note
    x = 0
    for st, lab in d['legend']:
        fc, ec, _ = ROAD_FILL[st]
        ax.add_patch(FancyBboxPatch((x, 0.962), 0.022, 0.022, boxstyle='round,pad=0,rounding_size=0.003', fc=fc, ec=ec,
                                    lw=1 if st != 'hold' else 0, ls='--' if st == 'try' else '-'))
        ax.text(x + 0.03, 0.973, lab, fontsize=7.5, color=INK, va='center'); x += 0.03 + 0.012 * len(lab) + 0.035
    ax.text(1, 0.973, d['note'], fontsize=7.5, color=MUTED, va='center', ha='right')

    # header: stage, timing
    hy, hh = 0.83, 0.115
    ax.text(lw_ / 2, hy + hh / 2, '시나리오', fontsize=8, color=MUTED, ha='center', va='center')
    ax.text(rx + rw / 2, hy + hh / 2, '사업 구조', fontsize=8, color=MUTED, ha='center', va='center')
    for i, (k, t, when) in enumerate(d['stages']):
        ax.add_patch(FancyBboxPatch((cx[i], hy), cw, hh, boxstyle='round,pad=0,rounding_size=0.006', fc='white', ec=GRID, lw=1))
        ax.add_patch(FancyBboxPatch((cx[i], hy + hh - 0.01), cw, 0.01, boxstyle='square,pad=0', fc=BLUE, ec='none'))
        ax.text(cx[i] + cw / 2, hy + hh - 0.02, f'{k} {br(t)}', fontsize=7.4, weight='bold', color=INK, ha='center', va='top', linespacing=1.25)
        ax.text(cx[i] + cw / 2, hy + 0.012, when, fontsize=6.8, color=MUTED, ha='center', va='bottom')
    # gate row
    gy, gh = 0.73, 0.085
    ax.text(lw_ / 2, gy + gh / 2, '다음 단계\n진입 조건', fontsize=7, color=MUTED, ha='center', va='center')
    for i, g in enumerate(d['gates']):
        ax.text(cx[i] + cw / 2, gy + gh / 2, br(g), fontsize=6.6, color=INK, ha='center', va='center', linespacing=1.25)
        ax.annotate('', xy=(cx[i + 1] + 0.004, gy + gh / 2), xytext=(cx[i] + cw - 0.004, gy + gh / 2),
                    arrowprops=dict(arrowstyle='-|>', color=MUTED, lw=0.8))
    ax.plot([0, 1], [gy - 0.008, gy - 0.008], color=GRID, lw=1)

    # scenario rows
    rh, rg, y = 0.15, 0.022, gy - 0.03
    for sc in d['scenarios']:
        y -= rh
        ax.text(0.006, y + rh * 0.62, sc['name'], fontsize=9, weight='bold', color=INK, va='center')
        ax.text(0.006, y + rh * 0.3, br(sc['cond']), fontsize=6.4, color=MUTED, va='center', linespacing=1.25)
        for i, (st, txt) in enumerate(sc['cells']):
            cell(cx[i], y, cw, rh, st, txt, fs=6.9)
            if i and st != 'hold' and sc['cells'][i - 1][0] != 'hold':
                ax.annotate('', xy=(cx[i] + 0.001, y + rh / 2), xytext=(cx[i] - gap - 0.001, y + rh / 2),
                            arrowprops=dict(arrowstyle='-|>', color=MUTED, lw=0.8, shrinkA=0, shrinkB=0))
        ax.text(rx + 0.006, y + rh / 2, br(sc['result']), fontsize=6.9, color=INK, va='center', linespacing=1.3)
        y -= rg
    # precedents row
    ax.plot([0, 1], [y + rg / 2, y + rg / 2], color=GRID, lw=1)
    ph = 0.11
    y -= ph - 0.005
    ax.text(lw_ / 2, y + ph / 2, '선례\n(확인된 회사)', fontsize=7, color=MUTED, ha='center', va='center', linespacing=1.25)
    for i, c in enumerate(d['cases']):
        ax.text(cx[i] + cw / 2, y + ph / 2, br(c), fontsize=6.6, color=INK if c != '확인 중' else MUTED,
                ha='center', va='center', linespacing=1.25)
    fig.savefig(OUT / name, facecolor=SURF, dpi=220, bbox_inches='tight', pad_inches=0.06); plt.close(fig)


for name, d in CHARTS.items():
    if d['kind'] == 'roadmap':
        roadmap(name, d)
    elif d['kind'] == 'map':
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
