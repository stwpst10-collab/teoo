"""ROOMOVE 보고서 차트(PNG). 색은 작업 기준 5장 검증 팔레트(파랑 #2a78d6·주황 #eb6834)만 쓴다.
    python 20-ROOMOVE/tools/charts.py <out_dir>
"""
import sys
from pathlib import Path
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager

OUT = Path(sys.argv[1]); OUT.mkdir(parents=True, exist_ok=True)
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False
BLUE, ORANGE, INK, MUTED, GRID = '#2a78d6', '#eb6834', '#1a1a1a', '#6b7280', '#e5e7eb'


def base(ax):
    for s in ('top', 'right', 'left'):
        ax.spines[s].set_visible(False)
    ax.spines['bottom'].set_color('#9ca3af')
    ax.tick_params(colors=MUTED, labelsize=9, length=0)
    ax.yaxis.set_visible(False)


# 1. 이사 피해구제 건수 (한국소비자원 2026.4.21)
fig, ax = plt.subplots(figsize=(6.4, 2.6), dpi=200)
yrs, v = ['2023', '2024', '2025'], [603, 785, 961]
b = ax.bar(yrs, v, color=BLUE, width=0.5)
for r, x in zip(b, v):
    ax.text(r.get_x() + r.get_width() / 2, x + 15, f'{x:,}건', ha='center', color=INK, fontsize=10, fontweight='bold')
base(ax); ax.set_ylim(0, 1100)
fig.tight_layout(); fig.savefig(OUT / 'moving_complaints.png'); plt.close(fig)

# 2. 고객 1명당 기대 수수료 범위 (당사 추정)
fig, ax = plt.subplots(figsize=(6.4, 2.8), dpi=200)
rows = [('이사+입주청소', 12, 21.5), ('+부분 시공 200만원', 18, 27.5), ('+인터넷 1건', 28, 37.5), ('이사·청소+공사 2,700만원', 93, 102.5)]
for i, (lab, lo, hi) in enumerate(rows[::-1]):
    ax.barh(i, hi - lo, left=lo, color=ORANGE if '공사' in lab else BLUE, height=0.5)
    ax.text(hi + 1.5, i, f'{lo:g}만~{hi:g}만원', va='center', fontsize=9, color=INK)
ax.set_yticks(range(len(rows))); ax.set_yticklabels([r[0] for r in rows[::-1]], fontsize=9, color=INK)
for s in ('top', 'right', 'left'):
    ax.spines[s].set_visible(False)
ax.spines['bottom'].set_color('#9ca3af'); ax.tick_params(colors=MUTED, labelsize=8, length=0)
ax.set_xlim(0, 125); ax.set_xlabel('고객 1명당 수수료(만원)', fontsize=8, color=MUTED)
fig.tight_layout(); fig.savefig(OUT / 'take_per_customer.png'); plt.close(fig)

# 3. 피해 유형 중 일정·현장 판단 관련 비중 (한국소비자원 인용 보도)
fig, ax = plt.subplots(figsize=(6.4, 2.6), dpi=200)
rows = [('인테리어: 계약 불이행\n(공사 기간 미준수 등, 2023~26상)', 25.2), ('가구 배송: 지연·미배송\n(2025)', 26.4),
        ('소규모 이사: 추가 비용 요구\n(2025)', 24.9), ('청소: 추가 비용 요구\n(2023~26.1Q)', 20.5)]
for i, (lab, x) in enumerate(rows[::-1]):
    ax.barh(i, x, color=BLUE, height=0.5)
    ax.text(x + 0.6, i, f'{x}%', va='center', fontsize=9, color=INK, fontweight='bold')
ax.set_yticks(range(len(rows))); ax.set_yticklabels([r[0] for r in rows[::-1]], fontsize=8, color=INK)
for s in ('top', 'right', 'left', 'bottom'):
    ax.spines[s].set_visible(False)
ax.xaxis.set_visible(False); ax.tick_params(length=0); ax.set_xlim(0, 32)
fig.tight_layout(); fig.savefig(OUT / 'schedule_share.png'); plt.close(fig)
print('charts ->', OUT)
