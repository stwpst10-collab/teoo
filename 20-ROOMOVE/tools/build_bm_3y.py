# -*- coding: utf-8 -*-
"""ROOMOVE 3개년 BM 엑셀 (2026-11-28)

- A안: IR 덱 v1 구조 그대로(커머스 74만 원 바구니), B안: 성립 조건을 넣은 제안
- 계산은 모두 수식. '가정' 시트의 파란 글씨를 바꾸면 손익이 따라 움직임
- 근거: 리서치/시뮬레이션/ROOMOVE 논리 점검·3개년 BM 2026-11-28.md

사용: python 20-ROOMOVE/tools/build_bm_3y.py <출력 xlsx>
"""
import sys
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

BLUE = Font(color='0000FF')
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=14)
GREY = Font(color='808080', size=9)
HDR_FILL = PatternFill('solid', fgColor='1F3864')
HDR_FONT = Font(bold=True, color='FFFFFF')
SEC_FILL = PatternFill('solid', fgColor='E7E6E6')
KEY_FILL = PatternFill('solid', fgColor='DDEBF7')
INPUT_FILL = PatternFill('solid', fgColor='F2F7FF')
THIN = Side(style='thin', color='BFBFBF')
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
FMT = {'pct': '0.0%', 'won': '#,##0', 'cnt': '#,##0', 'dec': '0.0', 'eok': '#,##0.00,,,"억"'}
# 억 표시는 원 단위 값을 /1e8 해서 따로 씀 → 'eok2'
FMT['eok2'] = '#,##0.00"억";[Red]-#,##0.00"억"'

R_PRICE = '수익모델·규제 검증 2026-11-25'
R_ENTRY = '진입점·단계별 수익 근거 2026-11-25'
R_FIN = 'build_finance_model.py 기본 시나리오'
R_BM = '논리 점검·3개년 BM 2026-11-28'

# ---------------- 공통 가정 (한 값)
COMMON = [
    ('v1', '첫 달 방문 수', '회/월', 'cnt', 3000, '가정', R_FIN),
    ('vg', '방문 수 월 성장률', '%/월', 'pct', 0.08, '가정', R_FIN),
    ('vcap', '월 방문 수 상한', '회/월', 'cnt', 40000, '가정', R_FIN),
    ('c1', '방문 → 회원가입', '%', 'pct', 0.06, '가정', R_FIN),
    ('c2', '회원가입 → 공간 계획 생성', '%', 'pct', 0.40, '가정', R_FIN),
    ('c3', '공간 계획 → 거래 프로젝트', '%', 'pct', 0.30, '가정', R_FIN),
    ('paid', '자체 유입 중 유료 광고 비중', '%', 'pct', 0.70, '가정', R_FIN),
    ('cpc', '유료 방문 1회당 광고비', '원', 'won', 2000, '가정', R_FIN + ' (검색광고 클릭당 수천 원~1~2만 원, 업체 자료)'),
    ('cl_use', '입주청소 이용률', '%', 'pct', 0.70, '가정', R_FIN),
    ('cl_gmv', '입주청소 1건 거래액(32평)', '원', 'won', 350000, '업체 자료', R_PRICE + ' 3-1 (30만~45만 원)'),
    ('cl_take', '입주청소 수수료율', '%', 'pct', 0.10, '가정', '청소 중개 20~35% 보도(2019)보다 낮게. 업체 인터뷰 필요'),
    ('pt_use', '부분시공 이용률', '%', 'pct', 0.30, '가정', R_FIN),
    ('pt_gmv', '부분시공 1건 거래액', '원', 'won', 4000000, '가정', R_FIN + ' (1,500만 원 미만 경미 공사)'),
    ('pt_take', '부분시공 수수료율', '%', 'pct', 0.03, '가정', '팀 초안 3%. 오늘의집은 정액 광고 월 20만~80만 원(2022)'),
    ('mv_use', '이사 업체 연결 이용률', '%', 'pct', 0.70, '가정', R_FIN),
    ('mv_fix', '이사 제휴 정액(건당, 성과 수수료 아님)', '원', 'won', 60000, '가정', '화물법상 운송주선 위험 → 성과 수수료 제외. 정액 허용 여부 법 검토 필요. ' + R_PRICE + ' 5-1'),
    ('leak', '직거래 이탈(서비스 수수료 중 못 받는 비율)', '%', 'pct', 0.10, '가정', R_FIN),
    ('buy', '커머스 구매 가구 비율', '%', 'pct', 0.40, '가정', '측정 필요(MVP 실거래)'),
    ('plan_fee', '계획 이용료(원하는 고객만)', '원', 'won', 99000, '가정', '선결제 테스트로 확인'),
    ('lv_pt_use', '거주 중 1건: 부분시공 이용률', '%', 'pct', 0.50, '가정', '27장 IMPROVE·CARE, 수요 별도 검증'),
    ('lv_pt_gmv', '거주 중 1건: 부분시공 거래액', '원', 'won', 3000000, '가정', '욕실 부분 보수 등'),
    ('lv_basket', '거주 중 1건: 커머스 구매 목록', '원', 'won', 500000, '가정', '책상·수납·조명 등'),
    ('p_fee', '제휴처 월 이용료', '원/월', 'won', 300000, '가정', 'Updater는 운영자 무료(업체 자료). 받을 수 있는지 측정'),
    ('p_cac', '제휴처 1곳 확보비', '원', 'won', 1000000, '가정', R_FIN),
    ('p_act', '제휴처 연중 평균 가동 비율', '%', 'pct', 0.5, '가정', '연중 순차 확보 가정'),
    ('mg_cap', '매니저 1명 월 처리 건수(자동화 0%)', '건/월', 'cnt', 20, '가정', '웨딩플래너 월 8~10건(인터뷰 1건). MVP 실측 1순위'),
    ('mg_cost', '매니저 1명 월 비용(부담 포함)', '원/월', 'won', 4200000, '가정', '급여 350만 원 × 1.2'),
    ('lv_load', '거주 중 1건의 매니저 부하(이사 1건=1)', '배', 'dec', 0.5, '가정', '작업 수가 적음'),
    ('var_rate', '결제·정산·분쟁 충당(매출 대비)', '%', 'pct', 0.05, '가정', R_FIN),
    ('var_case', '프로젝트 1건 운영 변동비(알림톡·통화)', '원', 'won', 3000, '가정', R_FIN),
    ('hc_cost', '고정 인력 1명 월 비용(부담 포함)', '원/월', 'won', 5500000, '가정', R_FIN),
    ('etc', '서버·법무·사무 등 월 고정비', '원/월', 'won', 4000000, '가정', R_FIN),
    ('dev', '초기 개발비(1년차 한 번)', '원', 'won', 70000000, '가정', R_FIN),
]

# ---------------- 연도별 가정: (키, 항목, 단위, fmt, A[3], B[3], 근거)
YEARLY = [
    ('ramp', '자체 유입 가동 비율(1년차 상반기 MVP)', '%', 'pct', [0.5, 1, 1], [0.5, 1, 1], '28장: 1단계 MVP 10~15가구 후 출시'),
    ('pshare', '제휴 채널 유입 비중(전체 프로젝트 중)', '%', 'pct', [0, 0, 0], [0, 0.3, 0.5], '24장 채널 후보(입주 커뮤니티·부동산 계약 직후·건설사). Updater식 B2B2C'),
    ('lv_ratio', '거주 중 프로젝트(이사 프로젝트 대비)', '%', 'pct', [0, 0, 0.3], [0, 0, 0.3], '27장 확장, 3년차부터'),
    ('basket', '커머스 구매 목록(구매 가구 1곳)', '원', 'won', [740000] * 3, [740000, 2000000, 3000000], '19장 예시 74만 원 / B는 가구·가전 포함 확대 가정'),
    ('ctake', '커머스 수수료율', '%', 'pct', [0.03, 0.12, 0.12], [0.03, 0.12, 0.12], '1년차 제휴 3%(아마존 공식표) → 입점 판매 12%(오늘의집 10~23%, 신뢰도 낮음)'),
    ('plan_pay', '계획 이용료 결제 비율', '%', 'pct', [0.1] * 3, [0.1, 0.3, 0.3], '가정, 선결제 테스트'),
    ('partners', '제휴처 수(연중 평균)', '곳', 'cnt', [0, 6, 24], [0, 6, 24], '가정'),
    ('auto', '자동화율(매니저 처리 건수 상승)', '%', 'pct', [0, 0.2, 0.4], [0, 0.2, 0.4], '21장: 반복 업무부터 자동화'),
    ('hc', '고정 인력(개발·기획·영업)', '명', 'cnt', [4, 5, 6], [4, 5, 6], R_FIN),
]

COLS = {'A': ['C', 'D', 'E'], 'B': ['F', 'G', 'H']}


def hdr(ws, r, vals, start=1):
    for i, v in enumerate(vals):
        c = ws.cell(row=r, column=start + i, value=v)
        c.fill, c.font, c.border = HDR_FILL, HDR_FONT, BOX
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


def build(out):
    wb = Workbook()
    # ---------------- 가정
    ga = wb.active
    ga.title = '가정'
    ga['A1'] = 'ROOMOVE 3개년 BM 가정'; ga['A1'].font = TITLE
    ga['A2'] = '파란 글씨 = 입력값(바꾸면 손익 시트가 따라 움직임). 근거 성격이 "가정"인 값은 모두 MVP에서 측정할 대상'; ga['A2'].font = GREY
    name = {}
    r = 4
    ga.cell(row=r, column=1, value='1. 공통 가정 (A·B 같은 값)').font = BOLD
    ga.cell(row=r, column=1).fill = SEC_FILL
    r += 1
    hdr(ga, r, ['항목', '단위', '값', '근거 성격', '근거·출처'])
    r += 1
    for k, lab, unit, f, v, kind, src in COMMON:
        ga.cell(row=r, column=1, value=lab)
        ga.cell(row=r, column=2, value=unit)
        c = ga.cell(row=r, column=3, value=v); c.font = BLUE; c.fill = INPUT_FILL; c.number_format = FMT[f]
        ga.cell(row=r, column=4, value=kind)
        ga.cell(row=r, column=5, value=src).alignment = WRAP
        for cc in range(1, 6): ga.cell(row=r, column=cc).border = BOX
        name[k] = f"가정!$C${r}"
        r += 1
    r += 1
    ga.cell(row=r, column=1, value='2. 연도별 가정').font = BOLD
    ga.cell(row=r, column=1).fill = SEC_FILL
    r += 1
    hdr(ga, r, ['항목', '단위', 'A안 1년차', 'A안 2년차', 'A안 3년차', 'B안 1년차', 'B안 2년차', 'B안 3년차', '근거·출처'])
    r += 1
    yname = {}
    for k, lab, unit, f, a, b, src in YEARLY:
        ga.cell(row=r, column=1, value=lab)
        ga.cell(row=r, column=2, value=unit)
        for i, v in enumerate(a + b):
            c = ga.cell(row=r, column=3 + i, value=v); c.font = BLUE; c.fill = INPUT_FILL; c.number_format = FMT[f]
        ga.cell(row=r, column=9, value=src).alignment = WRAP
        for cc in range(1, 10): ga.cell(row=r, column=cc).border = BOX
        yname[k] = r
        r += 1
    ga.column_dimensions['A'].width = 40
    ga.column_dimensions['B'].width = 8
    for col in 'CDEFGH': ga.column_dimensions[col].width = 13
    ga.column_dimensions['I'].width = 60
    # 공통 섹션의 근거·출처 칸(E)은 넓게: E도 13이라 줄바꿈 → 공통 블록은 E를 길게 쓰기 위해 별도 처리
    ga.column_dimensions['E'].width = 13
    for rr in range(6, 6 + len(COMMON)):
        ga.merge_cells(start_row=rr, start_column=5, end_row=rr, end_column=9)
    ga.cell(row=5, column=5).value = '근거·출처'
    ga.merge_cells(start_row=5, start_column=5, end_row=5, end_column=9)
    ga.freeze_panes = 'C4'

    # ---------------- 월별 방문
    mv = wb.create_sheet('월별방문')
    mv['A1'] = '월별 방문 수 = MIN(상한, 첫 달 × (1+성장률)^(월−1))'; mv['A1'].font = BOLD
    hdr(mv, 2, ['개월차', '연차', '방문 수'])
    for m in range(1, 37):
        rr = 2 + m
        mv.cell(row=rr, column=1, value=m)
        mv.cell(row=rr, column=2, value=f'=ROUNDUP(A{rr}/12,0)')
        c = mv.cell(row=rr, column=3, value=f"=MIN({name['vcap']},{name['v1']}*(1+{name['vg']})^(A{rr}-1))")
        c.number_format = FMT['cnt']
    mv.column_dimensions['C'].width = 12

    # ---------------- 손익
    pl = wb.create_sheet('손익')
    pl['A1'] = 'ROOMOVE 3개년 손익 (A안: 덱 v1 그대로 / B안: 성립 조건 반영)'; pl['A1'].font = TITLE
    pl['A2'] = '단위: 건·원. 억 표시 행은 원 ÷ 1억. 모든 값은 가정에서 계산한 당사 추정이며 MVP 결과로 다시 산정'; pl['A2'].font = GREY
    hdr(pl, 4, ['구분', '항목', 'A안 1년차', 'A안 2년차', 'A안 3년차', 'B안 1년차', 'B안 2년차', 'B안 3년차', '계산식'])
    N = name

    def Y(k, col):
        return f'가정!{col}${yname[k]}'

    rows = []  # (구분, 항목, fmt, formula(col, yi, rowref), 설명, key?)
    R = {}

    def add(sec_, key, lab, f, fn, desc, keyrow=False):
        rows.append((sec_, key, lab, f, fn, desc, keyrow))

    add('유입', 'visits', '방문 수(연)', 'cnt', lambda c, y: f'=SUMIF(월별방문!$B$3:$B$38,{y},월별방문!$C$3:$C$38)', '월별방문 시트 합')
    add('유입', 'own', '자체 유입 프로젝트', 'cnt', lambda c, y: f"={{visits}}*{N['c1']}*{N['c2']}*{N['c3']}*{Y('ramp', c)}", '방문 × 가입 × 계획 × 거래 × 가동 비율')
    add('유입', 'mv', '이사 프로젝트(제휴 채널 포함)', 'cnt', lambda c, y: f"={{own}}/(1-{Y('pshare', c)})", '자체 ÷ (1 − 제휴 비중)', True)
    add('유입', 'lv', '거주 중 프로젝트', 'cnt', lambda c, y: f"={{mv}}*{Y('lv_ratio', c)}", '이사 프로젝트 × 비율', True)
    add('건당 매출', 'svc', '이사 1건: 서비스 수수료(이탈 반영)', 'won',
        lambda c, y: f"=({N['cl_use']}*{N['cl_gmv']}*{N['cl_take']}+{N['pt_use']}*{N['pt_gmv']}*{N['pt_take']}+{N['mv_use']}*{N['mv_fix']})*(1-{N['leak']})",
        '(청소 + 부분시공 + 이사 정액) × (1 − 이탈)')
    add('건당 매출', 'com', '이사 1건: 커머스', 'won', lambda c, y: f"={N['buy']}*{Y('basket', c)}*{Y('ctake', c)}", '구매 비율 × 구매 목록 × 수수료율')
    add('건당 매출', 'pln', '이사 1건: 계획 이용료', 'won', lambda c, y: f"={Y('plan_pay', c)}*{N['plan_fee']}", '결제 비율 × 9.9만 원')
    add('건당 매출', 'rpp', '이사 1건 매출 합계', 'won', lambda c, y: '={svc}+{com}+{pln}', '', True)
    add('건당 매출', 'lrpp', '거주 중 1건 매출', 'won',
        lambda c, y: f"={N['lv_pt_use']}*{N['lv_pt_gmv']}*{N['pt_take']}*(1-{N['leak']})+{N['buy']}*{N['lv_basket']}*{Y('ctake', c)}", '부분시공 수수료 + 커머스')
    add('매출', 'rev_mv', '이사 프로젝트 매출', 'won', lambda c, y: '={mv}*{rpp}', '')
    add('매출', 'rev_lv', '거주 중 프로젝트 매출', 'won', lambda c, y: '={lv}*{lrpp}', '')
    add('매출', 'rev_p', '제휴처 이용료', 'won', lambda c, y: f"={Y('partners', c)}*{N['p_fee']}*12*{N['p_act']}", '제휴처 수 × 월 이용료 × 12 × 가동 비율')
    add('매출', 'rev', '매출 합계', 'won', lambda c, y: '={rev_mv}+{rev_lv}+{rev_p}', '', True)
    add('비용', 'ads', '광고비(자체 유입)', 'won', lambda c, y: f"={{visits}}*{Y('ramp', c)}*{N['paid']}*{N['cpc']}", '방문 × 가동 × 유료 비중 × 클릭당')
    add('비용', 'pcac', '제휴처 확보비', 'won', lambda c, y: f"={Y('partners', c)}*{N['p_cac']}", '')
    add('비용', 'mgn', '매니저 수', 'dec', lambda c, y: f"=MAX(1,({{mv}}+{{lv}}*{N['lv_load']})/12/({N['mg_cap']}/(1-{Y('auto', c)})))", '(이사 + 거주 중 × 부하) ÷ 12 ÷ (처리 건수 ÷ (1 − 자동화))')
    add('비용', 'mgc', '매니저 비용', 'won', lambda c, y: f"={{mgn}}*{N['mg_cost']}*12", '')
    add('비용', 'var', '변동비(결제·충당·운영)', 'won', lambda c, y: f"={{rev}}*{N['var_rate']}+({{mv}}+{{lv}})*{N['var_case']}", '')
    add('비용', 'fix', '고정비(인력·서버·초기 개발)', 'won', lambda c, y: f"={Y('hc', c)}*{N['hc_cost']}*12+{N['etc']}*12+IF({y}=1,{N['dev']},0)", '')
    add('비용', 'cost', '비용 합계', 'won', lambda c, y: '={ads}+{pcac}+{mgc}+{var}+{fix}', '', True)
    add('손익', 'op', '영업손익', 'won', lambda c, y: '={rev}-{cost}', '', True)
    add('손익', 'cum', '누적 영업손익', 'won', lambda c, y: '={op}' if y == 1 else '={op}+{cum@prev}', '')
    add('손익', 'rev_e', '매출(억)', 'eok2', lambda c, y: '={rev}/100000000', '', True)
    add('손익', 'op_e', '영업손익(억)', 'eok2', lambda c, y: '={op}/100000000', '', True)
    add('손익', 'cum_e', '누적 영업손익(억)', 'eok2', lambda c, y: '={cum}/100000000', '', True)
    add('단위경제', 'u_rev', '프로젝트 1건 평균 매출', 'won', lambda c, y: '=({rev_mv}+{rev_lv})/({mv}+{lv})', '제휴처 이용료 제외')
    add('단위경제', 'u_cac', '프로젝트 1건 평균 유입비', 'won', lambda c, y: '=({ads}+{pcac})/({mv}+{lv})', '')
    add('단위경제', 'u_mg', '프로젝트 1건 매니저비', 'won', lambda c, y: '={mgc}/({mv}+{lv})', '')
    add('단위경제', 'u_ct', '프로젝트 1건 공헌이익', 'won', lambda c, y: f"={{u_rev}}*(1-{N['var_rate']})-{N['var_case']}-{{u_cac}}-{{u_mg}}", '매출 − 변동비 − 유입비 − 매니저비. 28장 게이트: 양(+)', True)

    r = 5
    for sec_, key, lab, f, fn, desc, keyrow in rows:
        R[key] = r
        r += 1
    import re
    for sec_, key, lab, f, fn, desc, keyrow in rows:
        rr = R[key]
        pl.cell(row=rr, column=1, value=sec_)
        pl.cell(row=rr, column=2, value=lab)
        for sc in ('A', 'B'):
            for yi, col in enumerate(COLS[sc]):
                y = yi + 1
                formula = fn(col, y)
                prev = COLS[sc][yi - 1] if yi > 0 else None

                def sub(m):
                    k = m.group(1)
                    if k.endswith('@prev'):
                        return f'{prev}{R[k[:-5]]}'
                    return f'{col}{R[k]}'
                formula = re.sub(r'\{([a-z_]+(?:@prev)?)\}', sub, formula)
                c = pl.cell(row=rr, column=3 + (0 if sc == 'A' else 3) + yi, value=formula)
                c.number_format = FMT[f]
        pl.cell(row=rr, column=9, value=desc).font = GREY
        for cc in range(1, 10):
            cell = pl.cell(row=rr, column=cc)
            cell.border = BOX
            if keyrow:
                cell.fill = KEY_FILL
                if cc <= 8: cell.font = BOLD
    pl.column_dimensions['A'].width = 10
    pl.column_dimensions['B'].width = 34
    for col in 'CDEFGH': pl.column_dimensions[col].width = 15
    pl.column_dimensions['I'].width = 48
    pl.freeze_panes = 'C5'

    # 차트: 매출·영업손익(억)
    ch = BarChart()
    ch.type = 'col'
    ch.title = '매출과 영업손익(억 원): A안 1~3년차 | B안 1~3년차'
    ch.y_axis.title = '억 원'
    data = Reference(pl, min_col=2, max_col=8, min_row=R['rev_e'], max_row=R['op_e'])
    ch.add_data(data, from_rows=True, titles_from_data=True)
    ch.set_categories(Reference(pl, min_col=3, max_col=8, min_row=4, max_row=4))
    ch.height, ch.width = 9, 22
    pl.add_chart(ch, f'B{r + 2}')

    # ---------------- 측정할 숫자
    ms = wb.create_sheet('측정할 숫자')
    ms['A1'] = 'BM을 가르는 숫자 4개 (MVP에서 먼저 잴 것)'; ms['A1'].font = TITLE
    hdr(ms, 3, ['숫자', '지금 가정', '틀리면', '재는 방법'])
    meas = [
        ('프로젝트 1건 실제 거래액·구성', '덱 25장 759만 원(출처 없음) / 커머스 74만 원', 'BM 전체', 'MVP 10~15가구 실거래 기록'),
        ('매니저 1명 월 처리 건수', '20건(자동화 0%)', '매니저비가 매출을 넘음', 'MVP 건당 투입 시간 실측'),
        ('제휴 채널 유입 비중·유입비', 'B안 2년차 30%, 3년차 50%', '광고비가 매출을 넘음', '입주 커뮤니티·부동산 계약 직후 접점 실험(24장)'),
        ('계획 이용료 지불 의사', '10~30%', '고객 매출 0', '9.9만 원 선결제 테스트'),
    ]
    for i, row in enumerate(meas):
        for j, v in enumerate(row):
            c = ms.cell(row=4 + i, column=1 + j, value=v); c.alignment = WRAP; c.border = BOX
    for col, w in zip('ABCD', (30, 40, 26, 40)): ms.column_dimensions[col].width = w
    ms['A10'] = f'근거 문서: venture/20-ROOMOVE/리서치/시뮬레이션/ROOMOVE {R_BM}.md'; ms['A10'].font = GREY
    ms['A11'] = '이사 연결은 성과 수수료를 빼고 정액으로 둠(화물자동차 운수사업법 개정 2026-08-26 국회 통과, 무허가 유상 중개 금지). 정액 허용 여부도 법 검토 필요'; ms['A11'].font = GREY

    wb.move_sheet('손익', offset=-2)
    wb.active = 0
    wb.save(out)
    print('saved', out)


if __name__ == '__main__':
    build(sys.argv[1])
