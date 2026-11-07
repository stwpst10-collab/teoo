"""엑셀 시뮬레이션에 '해커톤 회의' 시트를 만든다(열려 있는 엑셀 창에 바로, 엑셀은 닫지 않음).

노란 칸(입력)만 바꾸면 3년 손익·단위당 이익·손익분기·TAM/SAM/SOM·민감도 표가 다시 계산된다.
B열 라벨·연차 머리(C~E열 1,2,3)는 기존 시트와 같아서 해커톤 문서 연동(excel_link.py)이 이 시트를 읽는다.

  python tools/pitch/excel_meeting.py   (기존 '해커톤 회의' 시트가 있으면 지우고 새로 만든다. 노란 칸 입력값은 유지)
"""
import sys
from pathlib import Path

import pythoncom
import win32com.client as win32

sys.path.insert(0, str(Path(__file__).resolve().parent))
import excel_link  # noqa: E402

SHEET = '해커톤 회의'
YELLOW, BLUE_HEAD, GREY = 0x99F2FF, 0xF7E8DA, 0xF2F2F2  # Excel은 BGR
NUM, WON, PCT = '#,##0;[빨강]-#,##0', '#,##0"원";[빨강]-#,##0"원"', '0.0%'

# (셀, 값, 서식)
INPUTS = [
    ('H3', '가정 입력 (노란 칸만 바꾸세요)', None),
    ('H4', '본체 판매가', None), ('I4', 99000, WON),
    ('H5', '본체 원가율', None), ('I5', 0.40, PCT),
    ('H6', '패치 판매가(개당)', None), ('I6', 8000, WON),
    ('H7', '패치 원가율', None), ('I7', 0.30, PCT),
    ('H8', '연 고정비', None), ('I8', 120000000, WON),
    ('H9', '판매·마케팅비(매출 대비)', None), ('I9', 0, PCT),
    ('H10', '연차', None), ('I10', 1, '0"년차"'), ('J10', 2, '0"년차"'), ('K10', 3, '0"년차"'),
    ('H11', '본체 판매 대수', None), ('I11', 1000, NUM), ('J11', 5000, NUM), ('K11', 10000, NUM),
    ('H12', '1대당 연간 패치 구매 수', None), ('I12', 3, NUM), ('J12', 6, NUM), ('K12', 6, NUM),
]
INPUT_CELLS = ['I4', 'I5', 'I6', 'I7', 'I8', 'I9', 'I11', 'J11', 'K11', 'I12', 'J12', 'K12']

# ---------- 시장 규모 가정: 연령대별 인구 × 심한 증상 비율 (H14부터) ----------
# 근거: 20-리서치/여성 인구와 월경 증상 통계 원표 확인.md, 월경 증상 유병률 기준과 대상 인구 보정.md
#  인구: 행정안전부 주민등록인구 2026.10 월말(KOSIS DT_1B04006 1세별·DT_1B04005N 5세별), 19세 = 19~49세 합 − 20~49세 합
#  2022 비율: 성·생식건강조사 2022 원 보고서 〈부표 3〉 10세 구간(청소년은 PHWR 그림 1A, 13~18세)
#  2025 비율: 같은 조사 2차(2025) 〈부표 3〉 19~39세·40~64세 2구간(15~18세 25.5%는 기존 시트 값, 13~18세)
#  현재 월경 중: 2022 〈부표 3〉 19~29 96.7%, 30~39 93.1%, 40~49 84.7%. 15~18세는 공표값 없음 → 100%
# (연령대, 기본 포함, 여성 인구, 적용 PMS 비율 기본값, 월경 중 비율, 참고 2022 PMS, 참고 2025 PMS, 참고 2022 월경통)
AGE_BANDS = [
    ('15~18세', 1, 909804, 0.255, 1.000, 0.293, 0.255, 0.426),
    ('19세', 1, 237246, 0.483, 0.967, 0.521, 0.483, 0.510),
    ('20~24세', 1, 1133735, 0.483, 0.967, 0.521, 0.483, 0.510),
    ('25~29세', 1, 1508529, 0.483, 0.967, 0.521, 0.483, 0.510),
    ('30~34세', 1, 1671535, 0.483, 0.931, 0.501, 0.483, 0.468),
    ('35~39세', 1, 1531550, 0.483, 0.931, 0.501, 0.483, 0.468),
    ('40~44세', 0, 1723637, 0.326, 0.847, 0.402, 0.326, 0.309),
    ('45~49세', 0, 1949288, 0.326, 0.847, 0.402, 0.326, 0.309),
]
MK_TOP = 14                      # 제목 행
MK_HEAD = MK_TOP + 1             # 열 머리
MK_FIRST = MK_HEAD + 1           # 첫 연령대 행
MK_LAST = MK_FIRST + len(AGE_BANDS) - 1
MK_SUM = MK_LAST + 1
MK_ADJ = MK_SUM + 2              # 월경 중 보정 적용 여부
MK_SAM = MK_ADJ + 1              # SAM 비율
MK_TAM_N = MK_SAM + 1            # TAM 인원(결과)
MK_SAM_N = MK_TAM_N + 1          # SAM 인원(결과)
MK_NOTE = MK_SAM_N + 2
MK_COLS = ['연령대', '포함(1/0)', '여성 인구(명)', '적용 심한 PMS 비율', '현재 월경 중 비율', '대상 인원(명)',
           '참고: 2022 PMS', '참고: 2025 PMS', '참고: 2022 월경통']
TAM_CELL, SAM_RATE_CELL = f'$I${MK_TAM_N}', f'$I${MK_SAM}'
MK_NOTES = [
    '인구: 행정안전부 주민등록인구 2026.10 월말(KOSIS DT_1B04006·DT_1B04005N). 19세 = 19~49세 합 − 20~49세 합.',
    '비율: 질병관리청 성·생식건강조사. 2022는 10세 구간(19~29·30~39·40~49), 2025는 19~39·40~64 두 구간만 공표. 심한 = VAS 7점 이상, 지난 3년 기준.',
    '15~18세 비율은 13~18세 조사값. 15~18세 현재 월경 중 비율은 공표값이 없어 100%로 둠.',
    '월경 중 보정 1 = 인구에 현재 월경 중 비율을 곱함(임신·무월경 제외, 15~39세 기준 TAM 약 5% 감소). 0 = 보정 없음(기존 방식).',
    'SAM 30%는 최종 보고서 v10.12 6.1의 중앙 설명값(범위 17~41%). 기본값은 기존 회의 시트와 같은 TAM(15~39세, 2025 비율, 보정 없음).',
]


def col_formulas(row_tpl):
    return [row_tpl.format(c=c, i=i) for c, i in zip('CDE', 'IJK')]


MODEL = [  # (행, B라벨, C~E 수식 템플릿, 서식)
    (2, '연차', None, '0"년차"'),
    (3, '본체', '=$I$4', NUM),
    (4, '수량', '={i}11', NUM),
    (5, '본체 매출', '={c}3*{c}4', NUM),
    (6, '본체 원가', '={c}5*$I$5', NUM),
    (9, '패치', '=$I$6', NUM),
    (10, '수량', '={c}4*{i}12', NUM),
    (11, '패치 매출', '={c}9*{c}10', NUM),
    (12, '패치 원가', '={c}11*$I$7', NUM),
    (15, '매출 총액', '={c}5+{c}11', NUM),
    (16, '원가총액', '={c}6+{c}12', NUM),
    (17, '매출이익', '={c}15-{c}16', NUM),
    (18, '매익률', '={c}17/{c}15', PCT),
    (19, '판매·마케팅비', '={c}15*$I$9', NUM),
    (20, '고정비', '=$I$8', NUM),
    (21, '영업이익', '={c}17-{c}19-{c}20', NUM),
    (22, '영업이익률', '={c}21/{c}15', PCT),
]
UNIT = [  # 단위당 이익(1년차 기준)
    (24, '단위당 이익(1년차 기준)', None, None),
    (25, '본체 1대 매출이익', '=$I$4*(1-$I$5)', NUM),
    (26, '패치 1개 매출이익', '=$I$6*(1-$I$7)', NUM),
    (27, '고객 1명 연간 매출이익', '=C25+C26*$I$12-($I$4+$I$6*$I$12)*$I$9', NUM),
    (28, '손익분기 본체 판매 대수', '=$I$8/C27', NUM),
]
TAM = [
    (30, '시장 규모', None, None),
    (32, '고객수', [f'={TAM_CELL}', f'=C32*{SAM_RATE_CELL}', '=$I$11'], NUM),
    (33, '시장규모', ['=C32*$I$4', '=D32*$I$4', '=E32*$I$4'], NUM),
    (34, 'SOM이 SAM에서 차지하는 비중', [None, None, '=E32/D32'], '0.00%'),
]
PRICES = [29900, 49000, 99000, 132000]
COSTS = [0.30, 0.40, 0.50, 0.60]


def market_block(ws):
    """연령대별 인구 × 심한 PMS 비율 표. 노란 칸(포함·인구·비율·월경 중 비율·보정·SAM)을 바꾸면 TAM/SAM이 다시 계산된다."""
    ws.Range(f'H{MK_TOP}').Value = '시장 규모 가정 — 연령대별 인구 × 심한 PMS 비율 (노란 칸 조정, 참고 열에서 골라 넣기)'
    ws.Range(f'H{MK_TOP}').Font.Bold = True
    head = ws.Range(f'H{MK_HEAD}:P{MK_HEAD}')
    head.Value = MK_COLS
    head.Interior.Color = BLUE_HEAD
    head.Font.Bold = True
    head.WrapText = True
    adj = f'$I${MK_ADJ}'
    for k, (band, inc, pop, rate, mens, r22, r25, d22) in enumerate(AGE_BANDS):
        r = MK_FIRST + k
        ws.Range(f'H{r}').Value = band
        for col, val, fmt in (('I', inc, '0'), ('J', pop, NUM), ('K', rate, PCT), ('L', mens, PCT),
                              ('N', r22, PCT), ('O', r25, PCT), ('P', d22, PCT)):
            ws.Range(f'{col}{r}').Value = val
            ws.Range(f'{col}{r}').NumberFormat = fmt
        ws.Range(f'M{r}').Formula = f'=I{r}*J{r}*K{r}*IF({adj}=1,L{r},1)'
        ws.Range(f'M{r}').NumberFormat = NUM
        ws.Range(f'N{r}:P{r}').Font.Color = 0x5C5F5F
    for c in 'IJKL':
        rng = ws.Range(f'{c}{MK_FIRST}:{c}{MK_LAST}')
        rng.Interior.Color = YELLOW
        rng.Font.Color = 0xC00000
    s = MK_SUM
    ws.Range(f'H{s}').Value = '합계(포함 연령대)'
    ws.Range(f'J{s}').Formula = f'=SUMPRODUCT(I{MK_FIRST}:I{MK_LAST},J{MK_FIRST}:J{MK_LAST})'
    ws.Range(f'K{s}').Formula = f'=IF(J{s}=0,0,M{s}/J{s})'  # 포함 인구 대비 평균 비율
    ws.Range(f'M{s}').Formula = f'=SUM(M{MK_FIRST}:M{MK_LAST})'
    ws.Range(f'J{s}').NumberFormat = NUM
    ws.Range(f'K{s}').NumberFormat = PCT
    ws.Range(f'M{s}').NumberFormat = NUM
    ws.Range(f'H{s}:M{s}').Font.Bold = True
    ws.Range(f'H{s}:P{s}').Interior.Color = GREY
    rows = [
        (MK_ADJ, '월경 중 보정 적용(1=적용)', 0, '0', True),
        (MK_SAM, 'SAM 비율(비약물 자가관리)', 0.30, PCT, True),
        (MK_TAM_N, 'TAM 대상 인원(명)', f'=M{MK_SUM}', NUM, False),
        (MK_SAM_N, 'SAM 대상 인원(명)', f'=I{MK_TAM_N}*I{MK_SAM}', NUM, False),
    ]
    for r, label, val, fmt, is_input in rows:
        ws.Range(f'H{r}').Value = label
        cell = ws.Range(f'I{r}')
        if isinstance(val, str):
            cell.Formula = val
        else:
            cell.Value = val
        cell.NumberFormat = fmt
        if is_input:
            cell.Interior.Color = YELLOW
            cell.Font.Color = 0xC00000
        else:
            ws.Range(f'H{r}:I{r}').Font.Bold = True
    ws.Range(f'J{MK_SAM}').Value = '범위 17~41%'
    ws.Range(f'J{MK_SAM}').Font.Color = 0x5C5F5F
    for k, txt in enumerate(MK_NOTES):
        c = ws.Range(f'H{MK_NOTE + k}')
        c.Value = txt
        c.Font.Size = 8
        c.Font.Color = 0x5C5F5F


def build():
    pythoncom.CoInitialize()
    path = excel_link.latest_workbook()
    try:
        xl = win32.GetActiveObject('Excel.Application')
    except pythoncom.com_error:
        sys.exit('엑셀이 열려 있지 않음')
    wb = None
    for i in range(1, xl.Workbooks.Count + 1):
        if xl.Workbooks(i).FullName.lower() == str(path).lower():
            wb = xl.Workbooks(i)
    if wb is None:
        sys.exit(f'열린 엑셀에서 {path.name}을 찾지 못함')
    xl.ScreenUpdating = False
    keep = {}  # 다시 만들 때 사용자가 바꾼 입력값 유지
    try:
        for i in range(1, wb.Worksheets.Count + 1):
            if wb.Worksheets(i).Name == SHEET:
                old = wb.Worksheets(i)
                cells = list(INPUT_CELLS)
                if str(old.Range(f'H{MK_TOP}').Value or '').startswith('시장 규모 가정 —'):  # 연령대 표가 있는 판
                    cells += [f'{c}{r}' for c in 'IJKL' for r in range(MK_FIRST, MK_LAST + 1)]
                    cells += [f'I{MK_ADJ}', f'I{MK_SAM}']
                keep = {c: old.Range(c).Value for c in cells if old.Range(c).Value is not None}
                xl.DisplayAlerts = False
                wb.Worksheets(i).Delete()
                xl.DisplayAlerts = True
                break
        ws = wb.Worksheets.Add(Before=wb.Worksheets(1))
        ws.Name = SHEET
        ws.Cells.Font.Name = '맑은 고딕'
        ws.Cells.Font.Size = 10
        ws.Range('B1').Value = '해커톤 회의용 시뮬레이션 — 오른쪽 노란 칸만 바꾸면 전체가 다시 계산됩니다'
        ws.Range('B1').Font.Bold = True
        ws.Range('B1').Font.Size = 12
        for cell, val, fmt in INPUTS:
            ws.Range(cell).Value = val
            if fmt:
                ws.Range(cell).NumberFormat = fmt
        for c in INPUT_CELLS:
            r = ws.Range(c)
            r.Interior.Color = YELLOW
            r.Font.Color = 0xC00000 >> 0  # 파란 글씨 = 입력값(금융 모델 관례)
        ws.Range('H3').Font.Bold = True
        market_block(ws)
        for c, v in keep.items():
            ws.Range(c).Value = v
        # 3년 손익
        ws.Range('C2').Value, ws.Range('D2').Value, ws.Range('E2').Value = 1, 2, 3
        for row, label, tpl, fmt in MODEL:
            ws.Range(f'B{row}').Value = label
            if tpl:
                for c, f in zip('CDE', col_formulas(tpl)):
                    ws.Range(f'{c}{row}').Formula = f
            ws.Range(f'C{row}:E{row}').NumberFormat = fmt
        for row in (2, 15, 17, 21):
            ws.Range(f'B{row}:E{row}').Font.Bold = True
        ws.Range('B2:E2').Interior.Color = BLUE_HEAD
        # 단위당 이익
        for row, label, f, fmt in UNIT:
            ws.Range(f'B{row}').Value = label
            if f:
                ws.Range(f'C{row}').Formula = f
                ws.Range(f'C{row}').NumberFormat = fmt
        ws.Range('B24').Font.Bold = True
        # TAM/SAM/SOM
        ws.Range('B30').Value = '시장 규모'
        ws.Range('B30').Font.Bold = True
        ws.Range('C31').Value, ws.Range('D31').Value, ws.Range('E31').Value = 'TAM(심한 PMS, 포함 연령대)', 'SAM(비약물 자가관리)', 'SOM(1년차 구매자)'
        ws.Range('C31:E31').Interior.Color = BLUE_HEAD
        for row, label, fs, fmt in TAM[1:]:
            ws.Range(f'B{row}').Value = label
            for c, f in zip('CDE', fs):
                if f:
                    ws.Range(f'{c}{row}').Formula = f
            ws.Range(f'C{row}:E{row}').NumberFormat = fmt
        # 민감도: 본체 가격 × 원가율 (1년차)
        top = 37
        for k, (title, metric) in enumerate((('1년차 매출이익률', 'gm'), ('1년차 영업이익(원)', 'op'))):
            r0 = top + k * 8
            ws.Range(f'B{r0}').Value = f'민감도: 본체 가격 × 본체 원가율 → {title} (수량·패치·고정비는 입력값)'
            ws.Range(f'B{r0}').Font.Bold = True
            ws.Range(f'B{r0 + 1}').Value = '본체 가격 \\ 원가율'
            for j, cst in enumerate(COSTS):
                cell = ws.Cells(r0 + 1, 3 + j)
                cell.Value = cst
                cell.NumberFormat = '0%'
                cell.Interior.Color = BLUE_HEAD
            for i, p in enumerate(PRICES):
                pc = ws.Cells(r0 + 2 + i, 2)
                pc.Value = p
                pc.NumberFormat = WON
                pc.Interior.Color = GREY
                for j in range(len(COSTS)):
                    P = f'$B{r0 + 2 + i}'
                    Cc = ws.Cells(r0 + 1, 3 + j).Address
                    gp = f'($I$11*{P}*(1-{Cc})+$I$11*$I$12*$I$6*(1-$I$7))'
                    rev = f'($I$11*{P}+$I$11*$I$12*$I$6)'
                    f = f'={gp}/{rev}' if metric == 'gm' else f'={gp}-{rev}*$I$9-$I$8'
                    c = ws.Cells(r0 + 2 + i, 3 + j)
                    c.Formula = f
                    c.NumberFormat = PCT if metric == 'gm' else NUM
        ws.Columns('B').ColumnWidth = 26
        for col in 'CDE':
            ws.Columns(col).ColumnWidth = 18
        ws.Columns('G').ColumnWidth = 3
        ws.Columns('H').ColumnWidth = 26
        for col in 'IJKLMNOP':
            ws.Columns(col).ColumnWidth = 14
        ws.Activate()
        ws.Range('I4').Select()
    finally:
        xl.ScreenUpdating = True
    wb.Save()
    v = {k: ws.Range(k).Value for k in ('C21', 'D21', 'E21', 'C28', 'C32', 'D32', 'E34', f'J{MK_SUM}')}
    print('만듦:', wb.Name, SHEET, v)


if __name__ == '__main__':
    build()
