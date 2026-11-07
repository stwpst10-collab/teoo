"""해커톤 보고서 차트. 각 차트는 ID로 등록하고 PNG 바이트를 돌려준다.

- excel=True 차트는 엑셀 시뮬레이션 값으로 그린다(sync 때마다 다시 그림).
- 나머지는 보고서 본문 수치(출처 번호 표기)로 그린다.
색은 기존 보고서와 같은 검증 팔레트 1·2번(파랑·주황)만 쓴다.
"""
import io

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

plt.rcParams.update({'font.family': 'Malgun Gothic', 'axes.unicode_minus': False, 'font.size': 10})
BLUE, ORANGE = '#2a78d6', '#eb6834'
BLUE_LIGHT = '#86b6ef'
INK, MUTED, GRID, SURF = '#0b0b0b', '#5f5f5c', '#e4e4e0', '#ffffff'
W = 6.4  # inch, 본문 폭에 맞춤

CHARTS = {}


def chart(cid, title, excel=False):
    def deco(fn):
        CHARTS[cid] = {'fn': fn, 'title': title, 'excel': excel}
        return fn
    return deco


def _frame(ax):
    ax.set_facecolor(SURF)
    for s in ('top', 'right', 'left'):
        ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_color(GRID)
    ax.tick_params(colors=MUTED, length=0)


def _png(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=220, facecolor=SURF, bbox_inches='tight')
    plt.close(fig)
    return buf.getvalue()


def _eok(v):
    return v / 1e8


@chart('pms_gap', '심한 PMS 경험률과 진료 비중')
def pms_gap(_=None):
    labels = ['19~39세 여성\n심한 PMS 경험률 [1]', '10~49세 여성 중\nPMS 진료 인원 비중 [2]']
    vals = [48.3, 0.14]
    fig, ax = plt.subplots(figsize=(W, 1.9))
    _frame(ax)
    y = [1, 0]
    ax.barh(y, vals, height=0.5, color=[BLUE, ORANGE])
    for yi, v in zip(y, vals):
        ax.text(v + 0.8, yi, f'{v:g}%', va='center', color=INK, fontsize=12, fontweight='bold')
    ax.set_yticks(y, labels, color=INK)
    ax.set_xlim(0, 60)
    ax.xaxis.set_visible(False)
    ax.text(0.99, -0.05, '2명 중 1명이 겪지만, 병원에 가는 사람은 1천명 중 1~2명', transform=ax.transAxes,
            ha='right', va='top', color=MUTED, fontsize=9)
    return _png(fig)


@chart('market_funnel', '시장 규모: 경험자에서 첫해 목표까지')
def market_funnel(_=None):
    steps = [
        ('19~49세 심한 PMS 경험 여성', 458, '429만~458만명'),
        ('핵심 타깃 25~39세', 239, '228만~239만명'),
        ('비약물 자가관리 경험층', 188, '73만~188만명'),
        ('첫해 판매 목표', 0.096, '960명 (범위 140~4,440명)'),
    ]
    fig, ax = plt.subplots(figsize=(W, 2.6))
    _frame(ax)
    y = list(range(len(steps)))[::-1]
    colors = [BLUE_LIGHT, BLUE, BLUE, ORANGE]
    for yi, (lab, v, txt), c in zip(y, steps, colors):
        ax.barh(yi, v, height=0.55, color=c)
        ax.text(max(v, 0) + 6, yi, txt, va='center', color=INK, fontsize=10)
    ax.set_yticks(y, [s[0] for s in steps], color=INK)
    ax.set_xlim(0, 640)
    ax.xaxis.set_visible(False)
    ax.set_title('상한값 기준 막대(만명). 첫해 목표는 점유율이 아닌 실행 목표', loc='left', fontsize=9, color=MUTED, pad=6)
    return _png(fig)


@chart('excel_pl_3y', '3년 매출 구성과 영업이익 (엑셀 연동)', excel=True)
def excel_pl_3y(d):
    r, yrs = d['rows'], d['years']
    xl = [f'{y}년차' for y in yrs]
    body, patch, op = [_eok(v) for v in r['본체 매출']], [_eok(v) for v in r['패치 매출']], [_eok(v) for v in r['영업이익']]
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(W, 2.8), gridspec_kw={'wspace': 0.35})
    for ax in (a1, a2):
        _frame(ax)
        ax.yaxis.grid(True, color=GRID, lw=0.6)
        ax.set_axisbelow(True)
    x = range(len(yrs))
    a1.bar(x, body, width=0.55, color=BLUE, label='본체', edgecolor=SURF, linewidth=1.5)
    a1.bar(x, patch, width=0.55, bottom=body, color=ORANGE, label='패치', edgecolor=SURF, linewidth=1.5)
    top = max(b + p for b, p in zip(body, patch))
    for xi, b, p in zip(x, body, patch):
        a1.text(xi, b + p + top * 0.03, f'{b + p:,.1f}억', ha='center', color=INK, fontsize=9.5)
    a1.set_xticks(list(x), xl, color=INK)
    a1.set_ylim(0, top * 1.18)
    a1.set_title('매출 (억원)', loc='left', fontsize=10, color=INK)
    a1.legend(frameon=False, fontsize=8.5, loc='upper left')
    a2.bar(x, op, width=0.55, color=[ORANGE if v < 0 else BLUE for v in op])
    a2.axhline(0, color=MUTED, lw=0.8)
    span = max(op) - min(min(op), 0)
    for xi, v in zip(x, op):
        a2.text(xi, v + (span * 0.04 if v >= 0 else -span * 0.04), f'{v:,.1f}억', ha='center',
                va='bottom' if v >= 0 else 'top', color=INK, fontsize=9.5)
    a2.set_xticks(list(x), xl, color=INK)
    a2.set_ylim(min(min(op), 0) - span * 0.15, max(op) + span * 0.15)
    a2.set_title('영업이익 (억원)', loc='left', fontsize=10, color=INK)
    return _png(fig)


def render(cid, excel_data=None):
    c = CHARTS[cid]
    if c['excel'] and excel_data is None:
        raise ValueError(f'{cid}는 엑셀 데이터가 필요함')
    return c['fn'](excel_data)
