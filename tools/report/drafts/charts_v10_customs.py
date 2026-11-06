"""Static report charts (validated palette slots 1-2; direct value labels for relief)."""
import sys
from pathlib import Path
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

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


def vbar(name, labels, values, unit, ymax, note=None):
    fig, ax = plt.subplots(figsize=(6.2, 2.7), dpi=220)
    fig.patch.set_facecolor(SURF); frame(ax)
    x = range(len(labels))
    ax.bar(list(x), values, width=0.5, color=BLUE)
    for xi, v in zip(x, values):
        ax.text(xi, v + ymax * 0.02, f'{v:,.1f}{unit}' if v % 1 else f'{v:,.0f}{unit}', ha='center', color=INK, fontsize=9)
    ax.set_xticks(list(x), labels, color=INK)
    ax.set_ylim(0, ymax); ax.yaxis.grid(True, color=GRID, lw=0.6); ax.set_axisbelow(True)
    ax.tick_params(axis='y', labelsize=8)
    if note:
        ax.set_title(note, loc='left', fontsize=8.5, color=MUTED, pad=6)
    fig.tight_layout(); fig.savefig(OUT / name, facecolor=SURF); plt.close(fig)


# 1. 심한 월경 증상 성인의 대처 (질병관리청 2022)
hbar('c1_coping.png',
     ['약국 진통제', '아무것도 안 함', '건강기능식품·민간요법', '한의원', '산부인과·비뇨의학과 진료'],
     [59.2, 23.5, 18.2, 4.8, 28.5], '%', 75,
     '심한 월경 증상(VAS 7점 이상)이 1개 이상인 성인 1,266명, 단위 %')

# 2. 심한 증상 유병률
hbar('c2_prevalence.png',
     ['성인 심한 월경전 증후군', '성인 심한 월경통', '청소년 심한 월경통', '청소년 심한 월경전 증후군'],
     [46.1, 41.6, 42.6, 29.3], '%', 60,
     '월경 중인 청소년 961명·성인 2,137명, VAS 7점 이상 비율, 단위 %')

# 3. 펨테크 시장 기관별 2024 추정
vbar('c3_femtech.png',
     ['Fortune Business\nInsights', 'Grand View\nResearch', 'Arizton\n(2023)', 'Global Market\nInsights'],
     [78.1, 392.9, 518.0, 602.0], '억$', 700,
     '글로벌 펨테크 시장 규모 추정(2024), 단위 억 달러 — 기관별 정의 상이')

# 4. 3개년 계획: 매출과 세전 손익 (같은 단위, 하나의 축)
fig, ax = plt.subplots(figsize=(6.2, 2.9), dpi=220)
fig.patch.set_facecolor(SURF); frame(ax)
yrs = ['1년차', '2년차', '3년차']; rev = [0.70, 3.00, 8.00]; pl = [-0.965, -0.44, 0.75]
w = 0.32; xs = range(3)
ax.bar([x - w / 2 - 0.01 for x in xs], rev, w, color=BLUE, label='매출')
ax.bar([x + w / 2 + 0.01 for x in xs], pl, w, color=ORANGE, label='세전 단순 손익')
for x, v in zip(xs, rev):
    ax.text(x - w / 2, v + 0.15, f'{v:.2f}', ha='center', color=INK, fontsize=8.5)
for x, v, t in zip(xs, pl, ['-0.965', '-0.44', '+0.75']):
    ax.text(x + w / 2, v + (0.15 if v >= 0 else -0.45), t, ha='center', color=INK, fontsize=8.5)
ax.axhline(0, color=MUTED, lw=0.8)
ax.set_xticks(list(xs), yrs, color=INK); ax.set_ylim(-1.6, 9)
ax.yaxis.grid(True, color=GRID, lw=0.6); ax.set_axisbelow(True); ax.tick_params(axis='y', labelsize=8)
ax.legend(frameon=False, loc='upper left', fontsize=8.5)
ax.set_title('린캔버스 계획 가정(실적 아님), 단위 억 원, 부가세 제외', loc='left', fontsize=8.5, color=MUTED, pad=6)
fig.tight_layout(); fig.savefig(OUT / 'c4_plan.png', facecolor=SURF); plt.close(fig)

# 5. 침투 시나리오
vbar('c5_scenario.png', ['침투 0.5%\n(약 2.35만 명)', '침투 1%\n(약 4.7만 명)', '침투 3%\n(약 14.1만 명)'],
     [28, 56, 169], '억 원', 200,
     '심한 월경통 응답 여성 470만 명(추정) × 누적 침투율 × 본체 12만 원 — 규모 예시, 연간 매출 아님')
print('charts ok')

# 6. 대형 홈헬스케어 2사 매출 추이 (같은 단위, 하나의 축)
fig, ax = plt.subplots(figsize=(6.2, 2.8), dpi=220)
fig.patch.set_facecolor(SURF); frame(ax)
yrs = ['2021', '2022', '2023', '2024', '2025']
cj = [6671, 7502, 5847, 5460, 5498]; bf = [5913, 5220, 4197, 4369, 4226]
w = 0.36; xs = range(5)
ax.bar([x - w / 2 - 0.01 for x in xs], cj, w, color=BLUE, label='세라젬')
ax.bar([x + w / 2 + 0.01 for x in xs], bf, w, color=ORANGE, label='바디프랜드')
for x, v in zip(xs, cj):
    ax.text(x - w / 2, v + 120, f'{v:,}', ha='center', color=INK, fontsize=7.5)
for x, v in zip(xs, bf):
    ax.text(x + w / 2, v + 120, f'{v:,}', ha='center', color=INK, fontsize=7.5)
ax.set_xticks(list(xs), yrs, color=INK); ax.set_ylim(0, 8800)
ax.yaxis.grid(True, color=GRID, lw=0.6); ax.set_axisbelow(True); ax.tick_params(axis='y', labelsize=8)
ax.legend(frameon=False, loc='upper right', fontsize=8.5, ncol=2)
ax.set_title('매출 추이, 단위 억 원', loc='left', fontsize=8.5, color=MUTED, pad=6)
fig.tight_layout(); fig.savefig(OUT / 'c6_homehealth.png', facecolor=SURF); plt.close(fig)

# 7. 국내 홈뷰티 디바이스 시장
vbar('c7_beauty.png', ['2018', '2022', '2030(전망)'], [0.5, 1.6, 3.4], '조 원', 4.0,
     '국내 홈뷰티 디바이스 시장 규모, 단위 조 원')
print('charts v3 ok')

# 8. 마사지용 기기(HSK 9019102000) 수입액 추이 — 관세청
vbar('c8_imports.png', ['2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'],
     [2.73, 3.80, 4.18, 4.05, 4.88, 5.65, 4.57, 3.70, 3.95, 3.49], '', 6.5,
     '마사지용 기기(HSK 9019102000) 수입액, 단위 억 달러')
print('chart c8 ok')
