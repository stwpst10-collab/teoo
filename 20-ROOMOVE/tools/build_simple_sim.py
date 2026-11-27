"""ROOMOVE 단순 시뮬레이션 v1: 방문 → 신청 → 계약 → 추가 이용 → 정기 서비스 깔때기 하나로 12개월을 계산.
파란 글씨·노란 칸만 바꾸면 전부 다시 계산된다. 실행: python 20-ROOMOVE/tools/build_simple_sim.py <출력.xlsx>"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment

OUT = sys.argv[1]
FN = '맑은 고딕'
BLUE = Font(name=FN, color='0000FF', size=10)
BLK = Font(name=FN, size=10)
BOLD = Font(name=FN, size=10, bold=True)
HEAD = Font(name=FN, size=10, bold=True, color='FFFFFF')
TITLE = Font(name=FN, size=14, bold=True)
GREEN = Font(name=FN, color='008000', size=10)
YEL = PatternFill('solid', fgColor='FFFF00')
HFILL = PatternFill('solid', fgColor='1F2937')
GRAY = PatternFill('solid', fgColor='F2F2F2')
thin = Side(style='thin', color='D0D0D0')
BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
WON = '#,##0;(#,##0);-'
PCT = '0.0%;(0.0%);-'
NUM = '#,##0.0;(#,##0.0);-'

wb = Workbook()

# ───────── 가정 ─────────
a = wb.active; a.title = '가정'
a['A1'] = 'ROOMOVE 단순 시뮬레이션: 가정'; a['A1'].font = TITLE
a['A2'] = '파란 글씨 = 바꿔 넣는 값(노란 칸은 먼저 측정할 핵심 값). 검은 글씨 = 계산. 값은 모두 측정 전 가정이며, 근거 성격 칸에 출처를 밝힘.'
a['A2'].font = Font(name=FN, size=9, color='555555')
hdr = ['구분', '항목', '값', '단위', '근거', '근거 성격', '측정 방법(→ 측정 계획 시트)']
for j, h in enumerate(hdr, 1):
    c = a.cell(row=4, column=j, value=h); c.font = HEAD; c.fill = HFILL; c.border = BOX
rows = [
    # (구분, 항목, 값, 단위, 서식, 근거, 성격, 측정, 핵심여부)
    ('유입', '월 방문자 수', 1000, '명/월', WON, '시작 규모 가정', '당사 가정', '①', False),
    ('유입', '방문 → 상담 신청 전환율', 0.05, '%', PCT, '근거 없음', '당사 가정', '①', True),
    ('유입', '상담 신청 → 이사 계약 전환율', 0.30, '%', PCT, '근거 없음', '당사 가정', '②', True),
    ('이사', '이사 1건 평균 거래액', 1000000, '원', WON, '포장이사 투룸 이상 평균 약 75만원, 5톤 95만~130만원의 중간값', '업체 자료(숨고·아정당)', '③', False),
    ('이사', '이사 수수료 방식(1=정률, 2=건당 정액)', 1, '선택', '0', '정률은 운송주선 허가가 있어야 함(보고서 3.8절)', '법령', '④', False),
    ('이사', '이사 정률 수수료율', 0.10, '%', PCT, '내부 논의 기준값, 시장 표준 아님', '당사 가정', '④', True),
    ('이사', '이사 건당 정액 수수료', 30000, '원/건', WON, '아정당 제휴 이사업체 가정이사 건당 3만원 예시', '업체 자료', '④', False),
    ('입주 직후', '추가 이용률(청소·설치·인터넷·수리)', 0.40, '%', PCT, '내부 논의 기준값', '당사 가정', '⑤', True),
    ('입주 직후', '추가 이용 1명 거래액', 300000, '원', WON, '내부 논의 기준값', '당사 가정', '⑤', False),
    ('입주 직후', '추가 이용 수수료율', 0.10, '%', PCT, '업체 수수료 가정(보고서 표 15와 같음)', '당사 가정', '④', False),
    ('거주 중', '정기 서비스 전환율(이사 고객 중)', 0.10, '%', PCT, '근거 없음. 생애가치를 가장 크게 바꾸는 값', '당사 가정', '⑥', True),
    ('거주 중', '정기 서비스 1명 월 거래액', '=66900*26/12', '원/월', WON, '미소 정기 청소 4시간 6만6,900원 × 2주 1회(연 26회) ÷ 12', '업체 자료 + 당사 가정(주기)', '⑥', False),
    ('거주 중', '정기 서비스 수수료율', 0.10, '%', PCT, '업체 수수료 가정', '당사 가정', '④', False),
    ('거주 중', '정기 서비스 월 유지율', 0.90, '%', PCT, '근거 없음(청소연구소 연 재이용 83%는 회사 발표, 정의 다름)', '당사 가정', '⑦', True),
    ('비용', '변동비율(이사매니저·결제 등, 수수료 대비)', 0.30, '%', PCT, '내부 논의 기준값', '당사 가정', '⑧', False),
    ('비용', '고객 확보 비용(이사 계약 1명당)', 70000, '원/명', WON, '내부 논의 기준값', '당사 가정', '⑨', True),
]
NAME = {}
r0 = 5
for i, (g, k, v, u, f, src, kind, ms, key) in enumerate(rows):
    r = r0 + i
    vals = [g, k, v, u, src, kind, ms]
    for j, x in enumerate(vals, 1):
        c = a.cell(row=r, column=j, value=x); c.font = BLK; c.border = BOX
        c.alignment = Alignment(vertical='center', wrap_text=(j in (2, 5, 6)))
    c = a.cell(row=r, column=3); c.font = BLUE if not str(v).startswith('=') else BLK; c.number_format = f
    if key: c.fill = YEL
    NAME[k] = f"가정!$C${r}"
for col, wdt in zip('ABCDEFG', [10, 34, 13, 8, 46, 22, 12]):
    a.column_dimensions[col].width = wdt
a.freeze_panes = 'A5'

def R(k): return NAME[k]

# 파생값
d0 = r0 + len(rows) + 1
a.cell(row=d0, column=1, value='계산된 값').font = BOLD
der = [
    ('월 이사 계약 고객', f"={R('월 방문자 수')}*{R('방문 → 상담 신청 전환율')}*{R('상담 신청 → 이사 계약 전환율')}", NUM, '명/월'),
    ('이사 1건 수수료', f"=IF({R('이사 수수료 방식(1=정률, 2=건당 정액)')}=1,{R('이사 1건 평균 거래액')}*{R('이사 정률 수수료율')},{R('이사 건당 정액 수수료')})", WON, '원'),
]
for i, (k, fml, f, u) in enumerate(der):
    r = d0 + 1 + i
    a.cell(row=r, column=2, value=k).font = BLK
    c = a.cell(row=r, column=3, value=fml); c.font = BLK; c.number_format = f
    a.cell(row=r, column=4, value=u).font = BLK
    NAME[k] = f"가정!$C${r}"

# ───────── 월별 계산 ─────────
m = wb.create_sheet('월별 계산')
m['A1'] = 'ROOMOVE 단순 시뮬레이션: 12개월(매달 같은 수의 이사 고객이 들어온다고 가정)'; m['A1'].font = TITLE
m.cell(row=3, column=1, value='항목').font = HEAD; m.cell(row=3, column=1).fill = HFILL
for j in range(12):
    c = m.cell(row=3, column=2 + j, value=f'{j + 1}개월'); c.font = HEAD; c.fill = HFILL; c.alignment = Alignment(horizontal='center')
c = m.cell(row=3, column=14, value='12개월 합계'); c.font = HEAD; c.fill = HFILL
from openpyxl.utils import get_column_letter as L
lines = [
    ('신규 이사 고객(명)', lambda col, prev: f"={R('월 이사 계약 고객')}", NUM, True),
    ('이사 수수료(원)', lambda col, prev: f"={col}4*{R('이사 1건 수수료')}", WON, True),
    ('추가 이용 고객(명)', lambda col, prev: f"={col}4*{R('추가 이용률(청소·설치·인터넷·수리)')}", NUM, True),
    ('추가 이용 수수료(원)', lambda col, prev: f"={col}6*{R('추가 이용 1명 거래액')}*{R('추가 이용 수수료율')}", WON, True),
    ('정기 서비스 신규(명)', lambda col, prev: f"={col}4*{R('정기 서비스 전환율(이사 고객 중)')}", NUM, True),
    ('정기 서비스 이용 고객(월말, 명)', lambda col, prev: (f"={col}8" if prev is None else f"={prev}9*{R('정기 서비스 월 유지율')}+{col}8"), NUM, False),
    ('정기 서비스 수수료(원)', lambda col, prev: f"={col}9*{R('정기 서비스 1명 월 거래액')}*{R('정기 서비스 수수료율')}", WON, True),
    ('매출 = 수수료 합계(원)', lambda col, prev: f"={col}5+{col}7+{col}10", WON, True),
    ('변동비(원)', lambda col, prev: f"={col}11*{R('변동비율(이사매니저·결제 등, 수수료 대비)')}", WON, True),
    ('공헌이익(원)', lambda col, prev: f"={col}11-{col}12", WON, True),
    ('고객 확보 비용(원)', lambda col, prev: f"={col}4*{R('고객 확보 비용(이사 계약 1명당)')}", WON, True),
    ('공헌이익 − 고객 확보 비용(원)', lambda col, prev: f"={col}13-{col}14", WON, True),
    ('누적(원)', lambda col, prev: (f"={col}15" if prev is None else f"={prev}16+{col}15"), WON, False),
]
for i, (lab, fn, f, summ) in enumerate(lines):
    r = 4 + i
    c = m.cell(row=r, column=1, value=lab); c.font = BOLD if r in (11, 13, 15, 16) else BLK
    for j in range(12):
        col = L(2 + j); prev = L(1 + j) if j > 0 else None
        cc = m.cell(row=r, column=2 + j, value=fn(col, prev)); cc.font = GREEN if r == 4 else BLK; cc.number_format = f
        if r in (11, 13, 15, 16): cc.fill = GRAY
    if summ:
        cc = m.cell(row=r, column=14, value=f"=SUM(B{r}:M{r})"); cc.font = BOLD; cc.number_format = f
m['N9'] = '=M9'; m['N9'].number_format = NUM; m['N16'] = '=M16'; m['N16'].number_format = WON
m['N9'].comment = Comment('12개월 말 이용 고객 수', 'ROOMOVE')
m.column_dimensions['A'].width = 32
for j in range(2, 15): m.column_dimensions[L(j)].width = 12
m.freeze_panes = 'B4'

# ───────── 고객 1명 단위경제 ─────────
u = wb.create_sheet('고객 1명')
u['A1'] = 'ROOMOVE 단순 시뮬레이션: 고객 1명 1년 단위경제'; u['A1'].font = TITLE
keep = f"(1-(1-{R('변동비율(이사매니저·결제 등, 수수료 대비)')}))"
cm = f"(1-{R('변동비율(이사매니저·결제 등, 수수료 대비)')})"
ret = R('정기 서비스 월 유지율')
items = [
    ('첫 이사 공헌이익(원)', f"={R('이사 1건 수수료')}*{cm}", WON),
    ('추가 이용 공헌이익(원, 기대값)', f"={R('추가 이용률(청소·설치·인터넷·수리)')}*{R('추가 이용 1명 거래액')}*{R('추가 이용 수수료율')}*{cm}", WON),
    ('정기 서비스 12개월 공헌이익(원, 기대값)', f"={R('정기 서비스 전환율(이사 고객 중)')}*{R('정기 서비스 1명 월 거래액')}*{R('정기 서비스 수수료율')}*{cm}*IF({ret}>=1,12,(1-{ret}^12)/(1-{ret}))", WON),
    ('1년 생애가치(원)', '=B4+B5+B6', WON),
    ('고객 확보 비용(원)', f"={R('고객 확보 비용(이사 계약 1명당)')}", WON),
    ('1년 생애가치 ÷ 고객 확보 비용(배)', '=IFERROR(B7/B8,0)', '0.00"배"'),
    ('첫 이사 공헌이익 ÷ 고객 확보 비용(배)', '=IFERROR(B4/B8,0)', '0.00"배"'),
    ('판단(보고서 4.3: 첫 이사만으로 2배 이상)', '=IF(B10>=2,"통과","미달")', '@'),
]
u.cell(row=3, column=1, value='항목').font = HEAD; u.cell(row=3, column=1).fill = HFILL
u.cell(row=3, column=2, value='값').font = HEAD; u.cell(row=3, column=2).fill = HFILL
for i, (k, fml, f) in enumerate(items):
    r = 4 + i
    u.cell(row=r, column=1, value=k).font = BOLD if r in (7, 9, 11) else BLK
    c = u.cell(row=r, column=2, value=fml); c.font = BOLD if r in (7, 9, 11) else BLK; c.number_format = f
u['A13'] = '정기 서비스 전환율에 따른 1년 생애가치 ÷ 고객 확보 비용(다른 값은 가정 시트 그대로)'; u['A13'].font = BOLD
u['A14'] = '정기 서비스 전환율'; u['B14'] = '배수'
for c in ('A14', 'B14'): u[c].font = HEAD; u[c].fill = HFILL
for i, rate in enumerate([0.0, 0.05, 0.10, 0.20, 0.30, 0.40]):
    r = 15 + i
    c = u.cell(row=r, column=1, value=rate); c.font = BLUE; c.number_format = PCT
    f = (f"=IFERROR((B$4+B$5+A{r}*{R('정기 서비스 1명 월 거래액')}*{R('정기 서비스 수수료율')}*{cm}"
         f"*IF({ret}>=1,12,(1-{ret}^12)/(1-{ret})))/B$8,0)")
    c = u.cell(row=r, column=2, value=f); c.font = BLK; c.number_format = '0.00"배"'
u['A22'] = '읽는 법: 첫 이사 공헌이익이 고객 확보 비용을 못 넘으면 정기 서비스 전환율이 사업성을 결정한다.'
u['A22'].font = Font(name=FN, size=9, color='555555')
u.column_dimensions['A'].width = 46; u.column_dimensions['B'].width = 16

# ───────── 측정 계획 ─────────
p = wb.create_sheet('측정 계획')
p['A1'] = 'ROOMOVE 단순 시뮬레이션: 전환율·수수료를 실제로 재는 방법'; p['A1'].font = TITLE
ph = ['번호', '무엇을', '어떻게 재나', '최소 표본', '바꿔 넣을 칸(가정 시트)']
for j, h in enumerate(ph, 1):
    c = p.cell(row=3, column=j, value=h); c.font = HEAD; c.fill = HFILL; c.border = BOX
plan = [
    ('①', '방문 → 상담 신청 전환율', '랜딩 페이지 1개 + 지역 커뮤니티·검색 소액 광고 2주. 방문자 중 "이사 일정 상담 신청" 남긴 비율', '방문 500명', '방문 → 상담 신청 전환율'),
    ('②', '상담 신청 → 이사 계약 전환율', '신청자에게 일정표·업체 연결 제안, 실제 이사 계약(업체 연결 완료)까지 간 비율', '신청 30명', '상담 신청 → 이사 계약 전환율'),
    ('③', '이사 1건 평균 거래액', '파일럿 계약서의 실제 금액 평균(평형·거리·포장 여부별로 기록)', '계약 20건', '이사 1건 평균 거래액'),
    ('④', '업체가 받아들이는 수수료율', '이사·청소·시공 업체 10곳 면담, 정률 5·10·15%와 건당 정액 3만원 중 수락하는 조건 기록(가상 계약서)', '업체 10곳', '이사 정률 수수료율 / 정액 / 추가 이용·정기 수수료율'),
    ('⑤', '입주 직후 추가 이용률', '이사 확정 고객에게 입주청소·설치·인터넷 묶음 제안, 입주 30일 안 이용 비율과 금액', '이사 고객 30명', '추가 이용률 / 추가 이용 1명 거래액'),
    ('⑥', '정기 서비스 전환율', '입주청소 뒤 정기 청소 첫 회 할인 제안, 60일 안 첫 정기 이용 비율', '이사 고객 30명', '정기 서비스 전환율'),
    ('⑦', '정기 서비스 월 유지율', '정기 이용자의 다음 달 재이용 여부를 3개월 추적', '정기 이용자 10명', '정기 서비스 월 유지율'),
    ('⑧', '변동비율', '이사매니저 1건당 투입 시간 × 시급 + 결제 수수료를 실측, 수수료 대비 비율', '계약 20건', '변동비율'),
    ('⑨', '고객 확보 비용', '채널별(광고·제휴) 쓴 돈 ÷ 그 채널로 계약한 고객 수', '채널당 계약 10명', '고객 확보 비용'),
]
for i, row in enumerate(plan):
    for j, x in enumerate(row, 1):
        c = p.cell(row=4 + i, column=j, value=x); c.font = BLK; c.border = BOX
        c.alignment = Alignment(vertical='center', wrap_text=True)
p['A14'] = '표본은 방향을 보기 위한 최소치(당사 제안)이며 통계적 확정치가 아님. 측정값이 나오면 가정 시트의 파란 칸만 바꾼다.'
p['A14'].font = Font(name=FN, size=9, color='555555')
for col, wdt in zip('ABCDE', [6, 26, 70, 14, 36]): p.column_dimensions[col].width = wdt

wb.save(OUT)
print('saved', OUT)
