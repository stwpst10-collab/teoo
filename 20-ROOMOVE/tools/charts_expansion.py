"""ROOMOVE 확장·경쟁 분석 보고서 차트(PNG). 색은 작업 기준 5장 검증 팔레트(파랑 #2a78d6·주황 #eb6834)만 쓴다.
    python 20-ROOMOVE/tools/charts_expansion.py <out_dir>
"""
import sys
from pathlib import Path
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

OUT = Path(sys.argv[1]); OUT.mkdir(parents=True, exist_ok=True)
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False
BLUE, ORANGE, INK, MUTED = '#2a78d6', '#eb6834', '#1a1a1a', '#6b7280'


def hbar_axes(ax):
    for s in ('top', 'right', 'left'):
        ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_color('#9ca3af')
    ax.tick_params(colors=MUTED, labelsize=8, length=0)


# 1. 미국 주택 구매 이사 가구의 1년 차 가구·가전·수리 지출 (NAHB 2026, 미국 소비자지출조사 2020~2023)
fig, ax = plt.subplots(figsize=(6.4, 2.4), dpi=200)
rows = [('신축 주택 구매 후 이사', 26882), ('기존 주택 구매 후 이사', 18673), ('이사하지 않은 자가 가구', 9457)]
for i, (lab, v) in enumerate(rows[::-1]):
    ax.barh(i, v, color=ORANGE if '않은' in lab else BLUE, height=0.5)
    ax.text(v + 400, i, f'{v:,}달러', va='center', fontsize=9, color=INK, fontweight='bold')
ax.set_yticks(range(len(rows))); ax.set_yticklabels([r[0] for r in rows[::-1]], fontsize=9, color=INK)
hbar_axes(ax); ax.set_xlim(0, 32000)
ax.set_xlabel('1년 차 가구·가전·수리 평균 지출(2023년 달러)', fontsize=8, color=MUTED)
fig.tight_layout(); fig.savefig(OUT / 'us_first_year_spend.png'); plt.close(fig)

# 2. 고객 1명 누적 수수료 범위 (당사 추정)
fig, ax = plt.subplots(figsize=(6.4, 2.4), dpi=200)
rows = [('첫 이사 1건(이사+입주청소)', 12, 21.5), ('5년 누적: 정기 청소 안 씀', 42, 62), ('5년 누적: 정기 청소 씀', 111, 134)]
for i, (lab, lo, hi) in enumerate(rows[::-1]):
    ax.barh(i, hi - lo, left=lo, color=ORANGE if '청소 씀' in lab else BLUE, height=0.5)
    ax.text(hi + 2, i, f'{lo:g}만~{hi:g}만원', va='center', fontsize=9, color=INK)
ax.set_yticks(range(len(rows))); ax.set_yticklabels([r[0] for r in rows[::-1]], fontsize=9, color=INK)
hbar_axes(ax); ax.set_xlim(0, 160)
ax.set_xlabel('고객 1명이 만드는 수수료(만원)', fontsize=8, color=MUTED)
fig.tight_layout(); fig.savefig(OUT / 'ltv_5y.png'); plt.close(fig)
print('charts ->', OUT)
