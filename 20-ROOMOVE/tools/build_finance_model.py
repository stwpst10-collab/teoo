# -*- coding: utf-8 -*-
"""ROOMOVE 재무 시뮬레이션 엑셀 빌더 (v1, 2026-11-26)

- 계산은 모두 엑셀 수식. 가정 시트의 파란 글씨(입력)를 바꾸면 전 시트가 따라 움직임
- 월별 계산 행은 ROWS 한 목록에서 두 번 그림
  1) 상세 시트(깔때기·매출·비용·손익·확장): 가정 시트 '적용값'(선택한 시나리오) 열 참조
  2) 시나리오계산 시트: 보수·기본·낙관 세 블록, 가정 시트 D·E·F 열을 각각 참조 → 요약에서 세 시나리오 동시 비교
  → 행을 더하거나 고칠 때 ROWS만 고치면 두 곳이 같이 바뀜(고도화 쉬운 구조)

사용: python 20-ROOMOVE/tools/build_finance_model.py <출력 xlsx 경로>
"""
import re
import sys
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation

MONTHS = 36
M0 = 8  # H열 = 1개월차
SCEN = ['보수', '기본', '낙관']
SCEN_COL = {'보수': 'D', '기본': 'E', '낙관': 'F', '선택': 'G'}

BLUE = Font(color='0000FF')
BLUE_B = Font(color='0000FF', bold=True)
BOLD = Font(bold=True)
GREY_F = Font(color='808080', size=9)
TITLE = Font(bold=True, size=14)
SEC_FILL = PatternFill('solid', fgColor='E7E6E6')
HDR_FILL = PatternFill('solid', fgColor='1F3864')
HDR_FONT = Font(bold=True, color='FFFFFF')
SEL_FILL = PatternFill('solid', fgColor='FFF2CC')
KEY_FILL = PatternFill('solid', fgColor='DDEBF7')
INPUT_FILL = PatternFill('solid', fgColor='F2F7FF')
THIN = Side(style='thin', color='BFBFBF')
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')

FMT = {
    'pct': '0.0%', 'won': '#,##0', 'cnt': '#,##0', 'cnt1': '#,##0.0', 'dec': '0.00',
    'mw': '#,##0.0,,;[Red]-#,##0.0,,', 'mon': '0"개월차"', 'txt': '@',
    'ppl': '#,##0', 'vis': '#,##0', 'gok': '#,##0', 'ppl1': '#,##0.0',
}
UNIT = {'pct': '%', 'won': '원', 'cnt': '건', 'cnt1': '건', 'mw': '백만원', 'mon': '', 'dec': '',
        'ppl': '명', 'vis': '회', 'gok': '곳', 'ppl1': '명'}

# ---------------------------------------------------------------- 가정
# (키, 항목, 단위, fmt, 보수, 기본, 낙관, 근거 성격, 근거·출처, 측정 필요, 팀 초안(비교용))
SECTIONS = []


def sec(title, rows):
    SECTIONS.append((title, rows))


R_PRICE = '리서치/수익모델·규제 검증 2026-11-25.md'
R_MKT = '리서치/시장 숫자 검증 2026-11-25.md'
R_ENTRY = '리서치/확장/진입점·단계별 수익 근거 2026-11-25.md'
R_INFO = '리서치/확장/고객정보 활용·이탈·재이용 2026-11-25.md'
R_SIM = '리서치/시뮬레이션/ROOMOVE 사업 아이디어 시뮬레이션 2026-11-25.md'

sec('A. 시장 (점유율 점검용, 매출 계산에는 안 씀)', [
    ('sam_resale', '수도권 아파트 매매 입주·부분 시공 가구(연)', '가구/년', 'cnt', 66000, 66000, 66000,
     '팀 추정(공식 자료 기반)', f'매매 72.6만(2025, 한국부동산원 인용) × 실제 입주 70% × 아파트 65% × 수도권 50% × 부분 시공 40%. 입주·시공 비율은 팀 추정. {R_MKT} 8절', '측정 필요', 'SAM 600억(수수료) 계산의 가구 수'),
    ('new_build', '수도권 신축 아파트 입주 가구(연)', '가구/년', 'cnt', 112064, 112064, 112064,
     '업체 자료', f'부동산R114 2026년 수도권 아파트 입주 11만2,064가구(2026.2 기준). {R_MKT} 7절', '', '-'),
])

sec('B. 유입·고객 깔때기 (방문 → 회원가입 → 실행계획 → 업체 계약)', [
    ('visit_m1', '첫 달 방문 수', '회/월', 'cnt', 2000, 3000, 5000,
     '팀 추정', '근거 없음. 다음 90일 과제 "목업 판매 페이지 방문·상담 전환율 측정"으로 확인', '측정 필요', '-'),
    ('visit_growth', '방문 수 월 성장률', '%/월', 'pct', 0.06, 0.08, 0.10,
     '팀 추정', '근거 없음', '측정 필요', '-'),
    ('visit_cap', '월 방문 수 상한', '회/월', 'cnt', 20000, 40000, 50000,
     '팀 추정', '깔때기 시트 "점유율" 행으로 점검: 보수는 3년차 연 관리 고객이 SOM(SAM의 1%) 수준이 되도록 둠', '측정 필요', '-'),
    ('paid_share', '유료 광고 유입 비중', '%', 'pct', 0.80, 0.70, 0.60,
     '팀 추정', '나머지는 입주박람회·카페·제휴 등 무상 유입으로 봄', '측정 필요', '-'),
    ('cpc', '유료 방문 1회당 광고비', '원', 'won', 3000, 2000, 1200,
     '업체 자료(참고)·팀 추정', f'"인테리어 업체" 검색광고 클릭당 수천 원~1~2만원(광고대행사 자료). 이사 소비자 키워드 단가는 확인 못 함. {R_PRICE} 1-4', '측정 필요', '-'),
    ('conv_signup', '방문 → 회원가입 전환율', '%', 'pct', 0.04, 0.06, 0.08,
     '팀 추정', '근거 없음', '측정 필요', '-'),
    ('conv_plan', '회원가입 → 이사 실행계획 생성 전환율', '%', 'pct', 0.30, 0.40, 0.50,
     '팀 추정', '근거 없음', '측정 필요', '-'),
    ('conv_manage', '실행계획 → 업체 계약(관리 고객) 전환율', '%', 'pct', 0.25, 0.30, 0.40,
     '팀 추정', f'관리 고객 = 업체 1곳 이상을 ROOMOVE로 계약해 매니저가 관리하는 이사. 숨고 견적→계약 전환율은 확인 못 함(견적비 역산 3~10% 가정). {R_PRICE} 1-1', '측정 필요', '-'),
])

sec('C. 업종별 이용률·거래액·수수료 (관리 고객 1건 기준)', [
    ('move_use', '이사: 관리 고객 중 이사 업체를 ROOMOVE로 계약', '%', 'pct', 0.60, 0.70, 0.80,
     '팀 추정', '근거 없음', '측정 필요', '-'),
    ('move_gmv', '이사: 1건 거래액', '원', 'won', 750000, 1000000, 1200000,
     '업체 자료', f'숨고 2026 25~30평대 약 75만, 30평대 포장이사 90만~170만(대표 120만). {R_MKT} 4절, {R_PRICE} 3-1', '', '-'),
    ('move_take', '이사: 성과 수수료율', '%', 'pct', 0.0, 0.0, 0.10,
     '공식 자료(법령)', f'이사 성사 수수료는 운송주선 해당 가능성, 2026.8 개정법은 허가 없는 유상 중개 금지. 보수·기본은 미탑재, 낙관만 운송주선(이사화물) 허가 취득 가정. {R_PRICE} 5-1', '국토부 질의', '10%'),
    ('move_fixed', '이사: 제휴 업체 건당 정액', '원', 'won', 0, 60000, 0,
     '팀 추정(업체 자료 기반)', f'기본 = 허가 이사업체 제휴 정액. 숨고식 견적 단가 3,000원 ÷ 계약 전환 5% = 계약 1건당 업체 고객 확보 비용 6만원 수준. 개정법상 정액 허용 여부 확인 필요. {R_PRICE} 1-1·5-1', '측정 필요', '-'),
    ('clean_use', '입주청소: 이용률', '%', 'pct', 0.60, 0.70, 0.80,
     '팀 추정', f'가중평균 예시의 청소 70%. {R_PRICE} 3-3', '측정 필요', '-'),
    ('clean_gmv', '입주청소: 1건 거래액', '원', 'won', 300000, 350000, 380000,
     '업체 자료·언론 보도', f'32평 30만~45만(대표 35만), 미소 34평 34만~42만. {R_PRICE} 3-1', '', '-'),
    ('clean_take', '입주청소: 성과 수수료율', '%', 'pct', 0.08, 0.10, 0.12,
     '언론 보도(참고)·팀 추정', f'청소 중개에서 업체·노동자 몫의 20~35%를 플랫폼이 가져간다는 보도(2019). 10%는 그보다 낮아 수용 여지. {R_PRICE} 1-6', '업체 인터뷰', '10%'),
    ('part_use', '부분 시공(줄눈·탄성코트·도배 등): 이용률', '%', 'pct', 0.20, 0.30, 0.40,
     '팀 추정', f'근거 없음. 가중평균 예시 30%, 시장 구조 40%. {R_PRICE} 3-3, {R_MKT} 8절', '측정 필요', '-'),
    ('part_gmv', '부분 시공: 1건 거래액', '원', 'won', 2000000, 4000000, 8000000,
     '팀 추정', f'근거 없음. 묶음 200만(수익모델 검증 3-3) ~ 800만(시장 숫자 검증 8절). 숨고 집 인테리어 평균 575만~1,328만(업체 자료)', '측정 필요', '-'),
    ('part_take', '부분 시공: 성과 수수료율', '%', 'pct', 0.02, 0.03, 0.03,
     '업체 자료(참고)·팀 추정', f'오늘의집 시공 광고 월 80만원으로 한 달 2건 계약하면 건당 약 1.5%(2022 자료). {R_PRICE} 1-2', '업체 인터뷰', '3%'),
    ('big_use', '큰 공사(1,500만원 이상): 이용률', '%', 'pct', 0.03, 0.05, 0.07,
     '팀 추정', f'가중평균 예시 공사 5%. 실내건축공사업 등록 업체만 연결(건설산업기본법). {R_PRICE} 3-3·5-2', '측정 필요', '-'),
    ('big_gmv', '큰 공사: 1건 거래액', '원', 'won', 27000000, 27000000, 27000000,
     '팀 초안 예시', '초안의 2,700만원 공사 예시를 그대로 씀', '측정 필요', '2,700만원 예시'),
    ('big_take', '큰 공사: 성과 수수료율', '%', 'pct', 0.02, 0.03, 0.03,
     '팀 추정', '부분 시공과 같은 논리', '업체 인터뷰', '3%'),
    ('net_use', '인터넷·렌탈: 가입률', '%', 'pct', 0.20, 0.30, 0.40,
     '업체 자료(참고)·팀 추정', f'미국 이사자 12개월 안 인터넷 가입 48%(Porch 2021 설문). ROOMOVE 경유 비율은 확인 못 함. {R_INFO} 4-3', '측정 필요', '-'),
    ('net_fixed', '인터넷·렌탈: ROOMOVE 몫(건당 정액)', '원', 'won', 50000, 80000, 100000,
     '공시·팀 추정', f'아정당 가입 1건 매출 약 55만·고객 사은품 약 30만 추정(2025 감사보고서 + 업체 주장 가입자 수). ROOMOVE 몫 10만원은 가정. {R_PRICE} 2-1', '제휴 문의', '건당 정액(금액 없음)'),
    ('coord_pay', '고객 유료 조율 이용료: 결제 비율', '%', 'pct', 0.0, 0.10, 0.20,
     '팀 추정', f'레드팀 보완안 "고객 선결제 테스트". {R_SIM} 7절', '측정 필요', '-'),
    ('coord_fee', '고객 유료 조율 이용료: 금액', '원', 'won', 99000, 99000, 99000,
     '팀 추정', f'시뮬레이션의 예시 금액 9.9만원. {R_SIM} 7절', '측정 필요', '-'),
    ('leakage', '직거래 이탈(업체 수수료 중 못 받는 비율)', '%', 'pct', 0.20, 0.10, 0.05,
     '학술 자료', f'청소 플랫폼 우회 83건 중 1건 이하(Astashkina 외 2023) ~ 우회 없었으면 이용 24% 증가(Karacaoglu 외 2022). {R_INFO} 3-1', '측정 필요', '-'),
])

sec('D. 이사매니저 (핵심 리스크: 규모가 커질 때 고정비가 어떻게 느는가)', [
    ('mgr_cap', '매니저 1명 월 처리 건수(자동화 0% 기준)', '건/월', 'cnt', 12, 20, 30,
     '언론 보도(참고)·팀 추정', f'웨딩플래너 월 8~10건(인터뷰 1건), 출시 게이트 월 15건, 동시 30건 목표. {R_PRICE} 3-4, {R_SIM} 3·7절', '측정 필요(수동 운영 5~10건)', '-'),
    ('auto_start', '자동화율: 출시 시점', '%', 'pct', 0.0, 0.0, 0.10,
     '팀 추정', '자동화율 = 순서표·요청서·확답 수집을 소프트웨어가 대신해 매니저 손이 안 가는 업무 비율. 실제 처리 건수 = 기준 건수 ÷ (1 − 자동화율)', '측정 필요', '-'),
    ('auto_end', '자동화율: 36개월차(출시부터 직선 증가)', '%', 'pct', 0.20, 0.40, 0.60,
     '팀 추정', '해외 사람 컨시어지(Moved·MoveEasy·Hello Alfred)는 소프트웨어·B2B로 전환하거나 피인수. 리서치/경쟁사 조사 2026-11-25.md', '측정 필요', '-'),
    ('mgr_salary', '매니저 월 급여', '원/월', 'won', 4000000, 3500000, 3000000,
     '팀 초안 범위', f'요청서 가정 월 300만~400만원. {R_PRICE} 3-4', '', '월 300만~400만원'),
    ('mgr_burden', '인건비 부담률(4대보험·퇴직금)', '%', 'pct', 0.20, 0.20, 0.20,
     '팀 추정', f'1.2배 가정. {R_PRICE} 3-4', '', '-'),
    ('mgr_min', '최소 매니저 수', '명', 'cnt', 1, 1, 1,
     '팀 추정', '고객이 적어도 1명은 둠', '', '-'),
    ('sub_load', '정기 서비스 고객 1명의 매니저 부하(관리 고객 1건 = 1)', '배', 'dec', 0.10, 0.10, 0.10,
     '팀 추정', '정기 청소는 앱 예약 중심이라 이사 관리보다 손이 덜 감', '측정 필요', '-'),
])

sec('E. 변동비', [
    ('pg_rate', '결제·정산 비용(매출 대비)', '%', 'pct', 0.03, 0.03, 0.03,
     '팀 추정', '근거 없음. 카드·전자결제 수수료 수준으로 둠', '', '-'),
    ('reserve', '분쟁·지연 보상 충당(매출 대비)', '%', 'pct', 0.03, 0.02, 0.01,
     '공식 자료(법령)·팀 추정', f'약관법 7조: 당사 고의·과실 책임은 뺄 수 없음. 오늘의집 시공 지연 하루 최대 20만원 보상 사례. {R_PRICE} 5-5', '', '-'),
    ('per_case', '관리 고객 1건당 운영 변동비(알림톡·통화·요청서)', '원', 'won', 5000, 3000, 2000,
     '팀 추정', '근거 없음', '', '-'),
])

sec('F. 고정비 (매니저 제외)', [
    ('fix_salary', '개발·기획·영업 인력 1명 월 인건비(부담 포함)', '원/월', 'won', 6000000, 5500000, 5000000,
     '팀 추정', '근거 없음', '', '-'),
    ('hc_y1', '고정 인력 1년차', '명', 'cnt', 5, 4, 4, '팀 추정', '', '', '-'),
    ('hc_y2', '고정 인력 2년차', '명', 'cnt', 6, 5, 5, '팀 추정', '', '', '-'),
    ('hc_y3', '고정 인력 3년차', '명', 'cnt', 7, 6, 6, '팀 추정', '', '', '-'),
    ('server', '서버·외부 도구(지도·알림톡 기본료 등)', '원/월', 'won', 3000000, 2000000, 1500000, '팀 추정', '', '', '-'),
    ('other_fixed', '사무·법무 검토·기타', '원/월', 'won', 3000000, 2000000, 2000000,
     '팀 추정', '이사 과금 법무 검토, 약관·개인정보 동의 설계 포함', '', '-'),
    ('init_dev', '초기 개발비(1개월차 한 번)', '원', 'won', 100000000, 70000000, 50000000, '팀 추정', '', '', '-'),
])

sec('G. 확장 단계 (켜고 끄기는 확장 시트)', [
    ('e1_start', '확장1 입주 직후 연계 판매: 시작 시점', '개월차', 'cnt', 19, 13, 7, '팀 추정', '확장·경쟁 분석 v1: 2~4단계는 1단계 고객이 쌓인 뒤 순서대로', '', '-'),
    ('e1_conv', '확장1: 관리 고객 중 연계 구매 전환율', '%', 'pct', 0.02, 0.05, 0.08,
     '업체 자료(참고)·팀 추정', f'통과 기준 제안 5%, 2% 미만이면 축소. Updater 한 주 구매율 8% 이상(외부 분석). {R_INFO} 4-3·5-2', '측정 필요', '-'),
    ('e1_basket', '확장1: 연계 구매액(1건)', '원', 'won', 1000000, 1500000, 2000000,
     '팀 추정', f'예시 150만원. {R_ENTRY} 3-3', '측정 필요', '-'),
    ('e1_take', '확장1: 제휴 수수료율', '%', 'pct', 0.02, 0.03, 0.04,
     '업체 자료', f'아마존 제휴 가구·주거 3%, 주방 4.5%, TV 2%. {R_ENTRY} 2-1', '', '-'),
    ('e2_start', '확장2 거주 중 정기 청소: 시작 시점', '개월차', 'cnt', 19, 13, 7, '팀 추정', '', '', '-'),
    ('e2_conv', '확장2: 관리 고객 중 정기 청소 전환율(입주 다음 달)', '%', 'pct', 0.03, 0.08, 0.15,
     '팀 추정', f'공개 자료 없음(첫 측정 대상). 통과 기준 제안: 12개월 재이용 30%. {R_ENTRY} 0절, {R_INFO} 5-2', '측정 필요', '-'),
    ('e2_churn', '확장2: 정기 고객 월 이탈률', '%/월', 'pct', 0.08, 0.04, 0.015,
     '업체 자료·학술 자료', f'청소연구소 연간 재이용 83%(≈월 1.5%, 2022 업체 발표), 일반 청소 1회 후 이탈 약 66%(Karacaoglu 외). {R_INFO} 4-1', '측정 필요', '-'),
    ('e2_price', '확장2: 정기 청소 1회 요금', '원', 'won', 59900, 59900, 59900, '업체 자료', f'미소 정기 4시간 5만9,900원. {R_ENTRY} 2-4', '', '-'),
    ('e2_freq', '확장2: 월 이용 횟수(2주 1회)', '회/월', 'dec', 2.17, 2.17, 2.17, '업체 자료', '26회 ÷ 12개월', '', '-'),
    ('e2_take', '확장2: 수수료율', '%', 'pct', 0.08, 0.10, 0.12, '팀 추정', f'초안의 청소 10%를 씀. {R_ENTRY} 3-3', '업체 인터뷰', '-'),
    ('e3_start', '확장3 제휴 채널(B2B2C): 시작 시점', '개월차', 'cnt', 19, 13, 10, '팀 추정', '', '', '-'),
    ('e3_new_partners', '확장3: 월 신규 제휴처 수(임대·분양 운영사, 중개·관리사무소, 입주민 앱)', '곳/월', 'cnt', 1, 2, 3, '팀 추정', '근거 없음', '측정 필요', '-'),
    ('e3_hh', '확장3: 제휴처 1곳의 월 이사 가구 수', '가구/월', 'cnt', 10, 20, 30, '팀 추정', '단지·운영사 규모에 따라 큼', '측정 필요', '-'),
    ('e3_plan_rate', '확장3: 초대 가구 → 실행계획 생성', '%', 'pct', 0.20, 0.30, 0.40,
     '업체 자료(참고)', f'Updater 초대 가구 30~40% 가입(외부 분석). {R_INFO} 핵심 결론 6', '측정 필요', '-'),
    ('e3_fee', '확장3: 제휴처 월 이용료(운영사용 소프트웨어)', '원/월', 'won', 0, 300000, 500000,
     '팀 추정', '미국 Moved는 임대 운영사 대상 입주 자동화 소프트웨어(B2B2C)로 전환(상위 조사). 국내 지불 의사는 확인 못 함', '측정 필요', '-'),
    ('e3_cac', '확장3: 제휴처 1곳 확보 영업비(한 번)', '원', 'won', 2000000, 1000000, 500000, '팀 추정', '근거 없음', '측정 필요', '-'),
    ('e4_start', '확장4 다음 이사 재이용: 시작 시점(재이사 시차 24개월 고정)', '개월차', 'cnt', 25, 25, 25,
     '공식 자료(참고)', '전세 2년 만기 기준. 임차 평균 거주 3.6년(2024 주거실태조사)', '', '-'),
    ('e4_rate', '확장4: 2년 안 다시 이사하는 비율', '%', 'pct', 0.20, 0.30, 0.30,
     '공식 자료(참고)', f'현 주택 거주 2년 이내 가구 30.3%(2024 주거실태조사). {R_INFO} 4-2', '', '-'),
    ('e4_reuse', '확장4: 다시 이사할 때 ROOMOVE 재이용률', '%', 'pct', 0.10, 0.25, 0.40, '팀 추정', '근거 없음. 확장 분석은 생애가치 계산에서 제외', '측정 필요', '-'),
])

TOGGLES = [  # (키, 단계, 기본값, 하는 일, 누가 돈을 내나, 근거 강도)
    ('e1', '확장1. 입주 직후 연계 판매', 1, '구매처와 관계없이 배송·설치 일정 관리, 가구·가전 연계 구매 제안', '판매처·설치 업체(제휴 수수료 3% 안팎)', '미국 지출 근거, 국내 전환율 없음'),
    ('e2', '확장2. 거주 중 정기 청소', 1, '입주청소 고객을 정기 청소로 연결(반복 매출)', '고객(정기 이용료), 업체(수수료)', '정기 청소 재이용 83~88%(업체 발표), 이사 고객 전환율 없음'),
    ('e3', '확장3. 제휴 채널(B2B2C)', 1, '임대·분양 운영사, 중개·관리사무소, 입주민 앱이 이사 가구를 초대', '제휴처(월 이용료), 업체(수수료)', '해외 성립 사례의 공통 조건(Updater·Moved), 국내 지불 의사 없음'),
    ('e4', '확장4. 다음 이사 재이용', 0, '저장된 가구 목록·평면으로 다시 이사 지원', '1단계와 같음', '장기 가능성, 확장 분석에서 생애가치 계산 제외 → 기본은 끔'),
]

# ---------------------------------------------------------------- 월별 계산 행
# (시트, 키, 항목, fmt, 합계 방식, 수식 템플릿)   키 None = 소제목
# 템플릿: {a:키} 가정, {t:키} 확장 켜기, {r:키} 같은 달, {p:키} 전월, {l24:키} 24개월 전, {m} 개월차
#         custom 행: {rng:키} 36개월 범위, {hdr} 개월차 머리글 범위
ROWS = [
    ('깔때기', None, '코어 유입: 방문 → 회원가입 → 이사 실행계획'),
    ('깔때기', 'visits', '방문 수', 'vis', 'sum', '=MIN({a:visit_cap},{a:visit_m1}*(1+{a:visit_growth})^({m}-1))'),
    ('깔때기', 'r_signup', '  전환율: 방문 → 회원가입', 'pct', 'end', '={a:conv_signup}'),
    ('깔때기', 'signups', '회원가입', 'cnt', 'sum', '={r:visits}*{r:r_signup}'),
    ('깔때기', 'r_plan', '  전환율: 회원가입 → 실행계획', 'pct', 'end', '={a:conv_plan}'),
    ('깔때기', 'plans_core', '이사 실행계획 생성(코어)', 'cnt', 'sum', '={r:signups}*{r:r_plan}'),
    ('깔때기', 'plans_e3', '이사 실행계획 생성(확장3 제휴 채널)', 'cnt', 'sum', '={r:e3_plans}'),
    ('깔때기', 'plans_total', '이사 실행계획 생성 합계', 'cnt', 'sum', '={r:plans_core}+{r:plans_e3}'),
    ('깔때기', None, '업체 계약 전환: 관리 고객(업체 1곳 이상 계약, 매니저가 관리)'),
    ('깔때기', 'r_manage', '  전환율: 실행계획 → 업체 계약', 'pct', 'end', '={a:conv_manage}'),
    ('깔때기', 'cust_core', '관리 고객(코어)', 'cnt', 'sum', '={r:plans_core}*{r:r_manage}'),
    ('깔때기', 'cust_e3', '관리 고객(확장3 제휴 채널)', 'cnt', 'sum', '={r:plans_e3}*{r:r_manage}'),
    ('깔때기', 'cust_e4', '관리 고객(확장4 재이용)', 'cnt', 'sum', '={r:e4_cust}'),
    ('깔때기', 'cust_total', '관리 고객 합계', 'cnt', 'sum', '={r:cust_core}+{r:cust_e3}+{r:cust_e4}'),
    ('깔때기', None, '업종별 계약 = 관리 고객 × 업종 이용률'),
    ('깔때기', 'r_move', '  이용률: 이사', 'pct', 'end', '={a:move_use}'),
    ('깔때기', 'n_move', '이사 계약', 'cnt', 'sum', '={r:cust_total}*{r:r_move}'),
    ('깔때기', 'r_clean', '  이용률: 입주청소', 'pct', 'end', '={a:clean_use}'),
    ('깔때기', 'n_clean', '입주청소 계약', 'cnt', 'sum', '={r:cust_total}*{r:r_clean}'),
    ('깔때기', 'r_part', '  이용률: 부분 시공', 'pct', 'end', '={a:part_use}'),
    ('깔때기', 'n_part', '부분 시공 계약', 'cnt', 'sum', '={r:cust_total}*{r:r_part}'),
    ('깔때기', 'r_big', '  이용률: 큰 공사', 'pct', 'end', '={a:big_use}'),
    ('깔때기', 'n_big', '큰 공사 계약', 'cnt1', 'sum', '={r:cust_total}*{r:r_big}'),
    ('깔때기', 'r_net', '  가입률: 인터넷·렌탈', 'pct', 'end', '={a:net_use}'),
    ('깔때기', 'n_net', '인터넷·렌탈 가입', 'cnt', 'sum', '={r:cust_total}*{r:r_net}'),
    ('깔때기', 'r_coord', '  결제 비율: 고객 조율 이용료', 'pct', 'end', '={a:coord_pay}'),
    ('깔때기', 'n_coord', '고객 조율 이용료 결제', 'cnt', 'sum', '={r:cust_total}*{r:r_coord}'),
    ('깔때기', None, '점검'),
    ('깔때기', 'share', '연환산 관리 고객 ÷ 수도권 대상 가구(매매 입주+신축 입주)', 'pct', 'end', '={r:cust_total}*12/({a:sam_resale}+{a:new_build})'),
    ('깔때기', 'cac', '관리 고객 1명 확보 비용(코어 광고비 ÷ 코어 관리 고객)', 'won', ('ratio', 'mkt_paid', 'cust_core'), '=IFERROR({r:mkt_paid}/{r:cust_core},0)'),

    ('매출', None, '업종별 거래액(GMV) × 수수료율 + 건당 정액'),
    ('매출', 'gmv_move', '이사 거래액', 'mw', 'sum', '={r:n_move}*{a:move_gmv}'),
    ('매출', 'rev_move', '  이사 매출(수수료 + 정액)', 'mw', 'sum', '={r:gmv_move}*{a:move_take}+{r:n_move}*{a:move_fixed}'),
    ('매출', 'gmv_clean', '입주청소 거래액', 'mw', 'sum', '={r:n_clean}*{a:clean_gmv}'),
    ('매출', 'rev_clean', '  입주청소 매출', 'mw', 'sum', '={r:gmv_clean}*{a:clean_take}'),
    ('매출', 'gmv_part', '부분 시공 거래액', 'mw', 'sum', '={r:n_part}*{a:part_gmv}'),
    ('매출', 'rev_part', '  부분 시공 매출', 'mw', 'sum', '={r:gmv_part}*{a:part_take}'),
    ('매출', 'gmv_big', '큰 공사 거래액', 'mw', 'sum', '={r:n_big}*{a:big_gmv}'),
    ('매출', 'rev_big', '  큰 공사 매출', 'mw', 'sum', '={r:gmv_big}*{a:big_take}'),
    ('매출', 'rev_net', '  인터넷·렌탈 매출(건당 정액)', 'mw', 'sum', '={r:n_net}*{a:net_fixed}'),
    ('매출', 'vendor_sub', '업체 수수료 소계', 'mw', 'sum', '={r:rev_move}+{r:rev_clean}+{r:rev_part}+{r:rev_big}+{r:rev_net}'),
    ('매출', 'leak', '  (−) 직거래 이탈', 'mw', 'sum', '=-{r:vendor_sub}*{a:leakage}'),
    ('매출', 'rev_coord', '  고객 조율 이용료', 'mw', 'sum', '={r:n_coord}*{a:coord_fee}'),
    ('매출', 'rev_core_svc', '이사 여정 서비스 순매출', 'mw', 'sum', '={r:vendor_sub}+{r:leak}+{r:rev_coord}'),
    ('매출', None, '확장 매출(확장 시트에서 켠 단계만)'),
    ('매출', 'rev_e1', '  확장1 연계 판매 수수료', 'mw', 'sum', '={r:e1_rev}'),
    ('매출', 'rev_e2', '  확장2 정기 청소 수수료', 'mw', 'sum', '={r:e2_rev}'),
    ('매출', 'rev_e3', '  확장3 제휴처 월 이용료', 'mw', 'sum', '={r:e3_fee_rev}'),
    ('매출', 'rev_total', '총매출', 'mw', 'sum', '={r:rev_core_svc}+{r:rev_e1}+{r:rev_e2}+{r:rev_e3}'),
    ('매출', 'gmv_total', '총거래액(GMV, 정률 수수료 대상)', 'mw', 'sum', '={r:gmv_move}+{r:gmv_clean}+{r:gmv_part}+{r:gmv_big}+{r:e1_gmv}+{r:e2_gmv}'),
    ('매출', 'take_eff', '실효 수수료율(총매출 ÷ 총거래액)', 'pct', ('ratio', 'rev_total', 'gmv_total'), '=IFERROR({r:rev_total}/{r:gmv_total},0)'),
    ('매출', 'rev_per_cust', '관리 고객 1건당 순매출(이사 여정 서비스)', 'won', ('ratio', 'rev_core_svc', 'cust_total'), '=IFERROR({r:rev_core_svc}/{r:cust_total},0)'),
    ('매출', 'rev_core_only', '코어 고객 매출(비교용, 확장 모두 끈 경우)', 'mw', 'sum', '=IFERROR({r:rev_core_svc}*{r:cust_core}/{r:cust_total},0)'),

    ('비용', None, '이사매니저: 처리 건수 → 필요 인원 → 인건비'),
    ('비용', 'auto', '자동화율', 'pct', 'end', '={a:auto_start}+({a:auto_end}-{a:auto_start})*({m}-1)/35'),
    ('비용', 'cap_eff', '매니저 1명 처리 가능 건수 = 기준 건수 ÷ (1 − 자동화율)', 'cnt1', 'end', '={a:mgr_cap}/(1-{r:auto})'),
    ('비용', 'workload', '매니저 업무량(관리 고객 + 정기 고객 × 부하)', 'cnt', 'sum', '={r:cust_total}+{r:e2_subs}*{a:sub_load}'),
    ('비용', 'mgrs', '필요 매니저 수', 'ppl', 'end', '=MAX({a:mgr_min},ROUNDUP({r:workload}/{r:cap_eff},0))'),
    ('비용', 'mgr_util', '매니저 1명 실제 처리 건수', 'cnt1', 'end', '={r:workload}/{r:mgrs}'),
    ('비용', 'mgr_cost', '매니저 인건비', 'mw', 'sum', '={r:mgrs}*{a:mgr_salary}*(1+{a:mgr_burden})'),
    ('비용', 'mgr_per_cust', '관리 고객 1건당 매니저 인건비', 'won', ('ratio', 'mgr_cost', 'cust_total'), '=IFERROR({r:mgr_cost}/{r:cust_total},0)'),
    ('비용', 'mgrs_core', '필요 매니저 수(코어만, 비교용)', 'ppl', 'end', '=MAX({a:mgr_min},ROUNDUP({r:cust_core}/{r:cap_eff},0))'),
    ('비용', 'mgr_cost_core', '매니저 인건비(코어만, 비교용)', 'mw', 'sum', '={r:mgrs_core}*{a:mgr_salary}*(1+{a:mgr_burden})'),
    ('비용', None, '마케팅(고객 확보 비용)'),
    ('비용', 'mkt_paid', '유료 광고비 = 방문 × 유료 비중 × 방문당 광고비', 'mw', 'sum', '={r:visits}*{a:paid_share}*{a:cpc}'),
    ('비용', 'mkt_e3', '제휴처 확보 영업비(확장3)', 'mw', 'sum', '={r:e3_cac_cost}'),
    ('비용', 'mkt_total', '마케팅 합계', 'mw', 'sum', '={r:mkt_paid}+{r:mkt_e3}'),
    ('비용', None, '결제·기타 변동비'),
    ('비용', 'var_pg', '결제·정산 비용', 'mw', 'sum', '={r:rev_total}*{a:pg_rate}'),
    ('비용', 'var_res', '분쟁·지연 보상 충당', 'mw', 'sum', '={r:rev_total}*{a:reserve}'),
    ('비용', 'var_case', '건당 운영 변동비', 'mw', 'sum', '={r:cust_total}*{a:per_case}'),
    ('비용', 'var_total', '변동비 합계', 'mw', 'sum', '={r:var_pg}+{r:var_res}+{r:var_case}'),
    ('비용', 'var_core', '변동비(코어만, 비교용)', 'mw', 'sum', '={r:rev_core_only}*({a:pg_rate}+{a:reserve})+{r:cust_core}*{a:per_case}'),
    ('비용', None, '개발·운영 고정비(매니저 제외)'),
    ('비용', 'hc', '고정 인력', 'ppl', 'end', '=CHOOSE(INT(({m}-1)/12)+1,{a:hc_y1},{a:hc_y2},{a:hc_y3})'),
    ('비용', 'fix_staff', '고정 인력 인건비', 'mw', 'sum', '={r:hc}*{a:fix_salary}'),
    ('비용', 'fix_server', '서버·외부 도구', 'mw', 'sum', '={a:server}'),
    ('비용', 'fix_other', '사무·법무·기타', 'mw', 'sum', '={a:other_fixed}'),
    ('비용', 'fix_init', '초기 개발비', 'mw', 'sum', '=IF({m}=1,{a:init_dev},0)'),
    ('비용', 'fix_total', '고정비 합계', 'mw', 'sum', '={r:fix_staff}+{r:fix_server}+{r:fix_other}+{r:fix_init}'),

    ('손익', None, '손익(코어 + 확장 시트에서 켠 단계)'),
    ('손익', 'pl_rev', '총매출', 'mw', 'sum', '={r:rev_total}'),
    ('손익', 'pl_var', '(−) 결제·기타 변동비', 'mw', 'sum', '=-{r:var_total}'),
    ('손익', 'pl_mgr', '(−) 이사매니저 인건비', 'mw', 'sum', '=-{r:mgr_cost}'),
    ('손익', 'pl_cm1', '공헌이익(마케팅 전)', 'mw', 'sum', '={r:pl_rev}+{r:pl_var}+{r:pl_mgr}'),
    ('손익', 'pl_mkt', '(−) 마케팅', 'mw', 'sum', '=-{r:mkt_total}'),
    ('손익', 'pl_cm2', '공헌이익(마케팅 후)', 'mw', 'sum', '={r:pl_cm1}+{r:pl_mkt}'),
    ('손익', 'pl_fix', '(−) 개발·운영 고정비', 'mw', 'sum', '=-{r:fix_total}'),
    ('손익', 'pl_op', '영업이익', 'mw', 'sum', '={r:pl_cm2}+{r:pl_fix}'),
    ('손익', 'pl_cum', '누적 영업이익', 'mw', 'end', '={p:pl_cum}+{r:pl_op}'),
    ('손익', 'pl_margin', '영업이익률', 'pct', ('ratio', 'pl_op', 'pl_rev'), '=IFERROR({r:pl_op}/{r:pl_rev},0)'),
    ('손익', None, '손익분기'),
    ('손익', 'be_month', '월 손익분기 시점(월 영업이익이 처음 0 이상인 달)', 'mon', 'custom',
     '=IFERROR(INDEX({hdr},MATCH(TRUE,INDEX({rng:pl_op}>=0,0),0)),"36개월 안 없음")'),
    ('손익', 'be_cum', '누적 손익 회복 시점(누적 영업이익이 처음 0 이상인 달)', 'mon', 'custom',
     '=IFERROR(INDEX({hdr},MATCH(TRUE,INDEX({rng:pl_cum}>=0,0),0)),"36개월 안 없음")'),
    ('손익', 'u_contrib', '관리 고객 1건당 공헌이익(마케팅 후, 이사 여정 서비스 기준)', 'won', 'end',
     '={r:rev_per_cust}*(1-{a:pg_rate}-{a:reserve})-{a:per_case}-{a:mgr_salary}*(1+{a:mgr_burden})/{r:cap_eff}-{a:paid_share}*{a:cpc}/({a:conv_signup}*{a:conv_plan}*{a:conv_manage})'),
    ('손익', 'be_cust', '손익분기 필요 관리 고객 수(그 달 고정비 ÷ 1건당 공헌이익, 초기 개발비 제외)', 'cnt', 'end',
     '=IF({r:u_contrib}>0,({r:fix_total}-{r:fix_init})/{r:u_contrib},"불가(공헌이익 음수)")'),
    ('손익', 'be_visits', '손익분기 필요 방문 수(필요 관리 고객 ÷ 방문→계약 전환율)', 'vis', 'end',
     '=IF(ISNUMBER({r:be_cust}),{r:be_cust}/({a:conv_signup}*{a:conv_plan}*{a:conv_manage}),"불가")'),
    ('손익', None, '비교: 코어만(확장 모두 끈 경우)'),
    ('손익', 'c_rev', '총매출(코어만)', 'mw', 'sum', '={r:rev_core_only}'),
    ('손익', 'c_var', '(−) 변동비', 'mw', 'sum', '=-{r:var_core}'),
    ('손익', 'c_mgr', '(−) 이사매니저 인건비', 'mw', 'sum', '=-{r:mgr_cost_core}'),
    ('손익', 'c_mkt', '(−) 마케팅(유료 광고)', 'mw', 'sum', '=-{r:mkt_paid}'),
    ('손익', 'c_fix', '(−) 개발·운영 고정비', 'mw', 'sum', '=-{r:fix_total}'),
    ('손익', 'c_op', '영업이익(코어만)', 'mw', 'sum', '={r:c_rev}+{r:c_var}+{r:c_mgr}+{r:c_mkt}+{r:c_fix}'),
    ('손익', 'c_cum', '누적 영업이익(코어만)', 'mw', 'end', '={p:c_cum}+{r:c_op}'),
    ('손익', 'c_be_month', '월 손익분기 시점(코어만)', 'mon', 'custom',
     '=IFERROR(INDEX({hdr},MATCH(TRUE,INDEX({rng:c_op}>=0,0),0)),"36개월 안 없음")'),

    ('확장', None, '확장1. 입주 직후 연계 판매(가구·가전·설치)'),
    ('확장', 'e1_on', '작동 여부(켜기 × 시작 시점 이후)', 'cnt', 'end', '=IF(AND({t:e1}=1,{m}>={a:e1_start}),1,0)'),
    ('확장', 'e1_n', '연계 구매 건수 = 관리 고객 × 전환율', 'cnt1', 'sum', '={r:e1_on}*{r:cust_total}*{a:e1_conv}'),
    ('확장', 'e1_gmv', '연계 구매 거래액', 'mw', 'sum', '={r:e1_n}*{a:e1_basket}'),
    ('확장', 'e1_rev', '연계 판매 수수료', 'mw', 'sum', '={r:e1_gmv}*{a:e1_take}'),
    ('확장', None, '확장2. 거주 중 정기 청소(반복 매출)'),
    ('확장', 'e2_on', '작동 여부', 'cnt', 'end', '=IF(AND({t:e2}=1,{m}>={a:e2_start}),1,0)'),
    ('확장', 'e2_new', '신규 정기 고객 = 전월 관리 고객 × 전환율', 'cnt1', 'sum', '={r:e2_on}*{p:cust_total}*{a:e2_conv}'),
    ('확장', 'e2_subs', '정기 고객(월말) = 전월 × (1 − 이탈률) + 신규', 'ppl', 'end', '={p:e2_subs}*(1-{a:e2_churn})+{r:e2_new}'),
    ('확장', 'e2_gmv', '정기 청소 거래액', 'mw', 'sum', '={r:e2_subs}*{a:e2_price}*{a:e2_freq}'),
    ('확장', 'e2_rev', '정기 청소 수수료', 'mw', 'sum', '={r:e2_gmv}*{a:e2_take}'),
    ('확장', None, '확장3. 제휴 채널(B2B2C: 임대·분양 운영사, 중개·관리사무소, 입주민 앱)'),
    ('확장', 'e3_on', '작동 여부', 'cnt', 'end', '=IF(AND({t:e3}=1,{m}>={a:e3_start}),1,0)'),
    ('확장', 'e3_newp', '신규 제휴처', 'gok', 'sum', '={r:e3_on}*{a:e3_new_partners}'),
    ('확장', 'e3_partners', '누적 제휴처', 'gok', 'end', '={p:e3_partners}+{r:e3_newp}'),
    ('확장', 'e3_invites', '제휴처가 초대한 이사 가구', 'cnt', 'sum', '={r:e3_partners}*{a:e3_hh}'),
    ('확장', 'e3_plans', '실행계획 생성 = 초대 × 생성률(→ 깔때기로 넘어감)', 'cnt', 'sum', '={r:e3_invites}*{a:e3_plan_rate}'),
    ('확장', 'e3_fee_rev', '제휴처 월 이용료 매출', 'mw', 'sum', '={r:e3_partners}*{a:e3_fee}'),
    ('확장', 'e3_cac_cost', '제휴처 확보 영업비', 'mw', 'sum', '={r:e3_newp}*{a:e3_cac}'),
    ('확장', None, '확장4. 다음 이사 재이용(24개월 전 관리 고객 기준)'),
    ('확장', 'e4_on', '작동 여부', 'cnt', 'end', '=IF(AND({t:e4}=1,{m}>={a:e4_start}),1,0)'),
    ('확장', 'e4_cust', '재이용 관리 고객 = 24개월 전 관리 고객 × 재이사 비율 × 재이용률(→ 깔때기)', 'cnt1', 'sum', '={r:e4_on}*{l24:cust_total}*{a:e4_rate}*{a:e4_reuse}'),
]

DETAIL_SHEETS = ['깔때기', '매출', '비용', '손익', '확장']
TOK = re.compile(r'\{(\w+)(?::(\w+))?\}')


class Model:
    def __init__(self, wb):
        self.wb = wb
        self.arow = {}      # 가정 키 → 행
        self.trow = {}      # 확장 켜기 키 → 행
        self.loc = {}       # (ctx, 키) → (시트, 행)
        self.hdr = {}       # (ctx, 시트) → 머리글 행

    # 셀 참조
    def ref(self, ctx, sheet_now, key, col):
        sh, r = self.loc[(ctx, key)]
        p = '' if sh == sheet_now else f"'{sh}'!"
        return f'{p}{col}{r}'

    def render(self, tpl, ctx, sheet_now, m):
        scol = '$' + SCEN_COL[ctx]

        def sub(mo):
            k, arg = mo.group(1), mo.group(2)
            if k == 'a':
                return f"'가정'!{scol}${self.arow[arg]}"
            if k == 't':
                return f"'확장'!$C${self.trow[arg]}"
            if k == 'm':
                return str(m)
            if k == 'r':
                return self.ref(ctx, sheet_now, arg, CL(M0 + m - 1))
            if k == 'p':
                return '0' if m == 1 else self.ref(ctx, sheet_now, arg, CL(M0 + m - 2))
            if k == 'l24':
                return '0' if m <= 24 else self.ref(ctx, sheet_now, arg, CL(M0 + m - 25))
            if k == 'rng':
                sh, r = self.loc[(ctx, arg)]
                p = '' if sh == sheet_now else f"'{sh}'!"
                return f'{p}$H${r}:${CL(M0 + MONTHS - 1)}${r}'
            if k == 'hdr':
                r = self.hdr[(ctx, sheet_now)] + 1
                return f'$H${r}:${CL(M0 + MONTHS - 1)}${r}'
            raise KeyError(k)
        return TOK.sub(sub, tpl)


def month_header(ws, r, title_cols=True):
    """r: 연차 행, r+1: 개월차 행"""
    labels = ['항목', '단위', '1년차', '2년차', '3년차', '3년 합계']
    for i, t in enumerate(labels):
        c = ws.cell(r + 1, i + 1, t)
        c.font, c.fill = HDR_FONT, HDR_FILL
    for m in range(1, MONTHS + 1):
        c1 = ws.cell(r, M0 + m - 1, (m - 1) // 12 + 1)
        c1.number_format = '0"년차"'
        c1.font = GREY_F
        c = ws.cell(r + 1, M0 + m - 1, m)
        c.number_format = '0"개월"'
        c.font, c.fill = HDR_FONT, HDR_FILL
        c.alignment = Alignment(horizontal='center')
    ws.cell(r + 1, 7).fill = HDR_FILL


KEY_ROWS = {'plans_total', 'cust_total', 'rev_core_svc', 'rev_total', 'mgrs', 'mgr_cost', 'mkt_total',
            'var_total', 'fix_total', 'pl_cm1', 'pl_cm2', 'pl_op', 'pl_cum', 'c_op', 'be_month', 'be_cum',
            'c_be_month', 'e1_rev', 'e2_rev', 'e3_plans', 'e4_cust'}


def layout(model, ctx, sheet_of, start_rows):
    """행 번호 배정. sheet_of(행 시트) → 실제로 그릴 시트명"""
    pos = dict(start_rows)
    for sh in set(sheet_of(s) for s in DETAIL_SHEETS):
        model.hdr[(ctx, sh)] = pos[sh]
        pos[sh] += 2
    for row in ROWS:
        sh = sheet_of(row[0])
        if row[1] is not None:
            model.loc[(ctx, row[1])] = (sh, pos[sh])
        pos[sh] += 1
    return pos


def write_rows(model, ctx, sheet_of):
    wb = model.wb
    done_hdr = set()
    pos_sec = {}
    for row in ROWS:
        shn = sheet_of(row[0])
        ws = wb[shn]
        if shn not in done_hdr:
            month_header(ws, model.hdr[(ctx, shn)])
            done_hdr.add(shn)
            pos_sec[shn] = model.hdr[(ctx, shn)] + 2
        r = pos_sec[shn]
        pos_sec[shn] += 1
        if row[1] is None:
            c = ws.cell(r, 1, row[2])
            c.font = BOLD
            for cc in range(1, M0 + MONTHS):
                ws.cell(r, cc).fill = SEC_FILL
            continue
        _, key, label, fmt, agg, tpl = row
        ws.cell(r, 1, label)
        ws.cell(r, 2, UNIT.get(fmt, ''))
        nf = FMT[fmt]
        if agg == 'custom':
            c = ws.cell(r, 3, model.render(tpl, ctx, shn, 1))
            c.number_format = nf
            c.font = BOLD
            c.fill = KEY_FILL
        else:
            for m in range(1, MONTHS + 1):
                c = ws.cell(r, M0 + m - 1, model.render(tpl, ctx, shn, m))
                c.number_format = nf
            yr = [(CL(M0), CL(M0 + 11)), (CL(M0 + 12), CL(M0 + 23)), (CL(M0 + 24), CL(M0 + 35))]
            for i, (a, b) in enumerate(yr):
                col = 3 + i
                if agg == 'sum':
                    f = f'=SUM({a}{r}:{b}{r})'
                elif agg == 'end':
                    f = f'={b}{r}'
                elif isinstance(agg, tuple):
                    num = model.ref(ctx, shn, agg[1], CL(col))
                    den = model.ref(ctx, shn, agg[2], CL(col))
                    f = f'=IFERROR({num}/{den},0)'
                ws.cell(r, col, f).number_format = nf
            if agg == 'sum':
                f = f'=SUM(C{r}:E{r})'
            elif agg == 'end':
                f = f'=E{r}'
            else:
                num = model.ref(ctx, shn, agg[1], 'F')
                den = model.ref(ctx, shn, agg[2], 'F')
                f = f'=IFERROR({num}/{den},0)'
            ws.cell(r, 6, f).number_format = nf
        if key in KEY_ROWS:
            for cc in range(1, M0 + MONTHS):
                ws.cell(r, cc).font = BOLD
            ws.cell(r, 1).fill = KEY_FILL


def set_widths(ws, first=46):
    ws.column_dimensions['A'].width = first
    ws.column_dimensions['B'].width = 8
    for c in 'CDEF':
        ws.column_dimensions[c].width = 11
    ws.column_dimensions['G'].width = 2
    for m in range(MONTHS):
        ws.column_dimensions[CL(M0 + m)].width = 8.5


# ---------------------------------------------------------------- 시트 작성
def build(out):
    wb = Workbook()
    names = ['안내', '요약', '가정', '깔때기', '매출', '비용', '손익', '확장', '민감도', '시나리오계산']
    wb.active.title = names[0]
    for n in names[1:]:
        wb.create_sheet(n)
    model = Model(wb)

    # ---- 가정
    ws = wb['가정']
    ws['A1'] = 'ROOMOVE 재무 시뮬레이션 v1 — 가정(입력)'
    ws['A1'].font = TITLE
    ws['A2'] = '파란 글씨만 입력값. 노란 칸에서 시나리오를 고르면 상세 시트(깔때기·매출·비용·손익·확장)가 그 시나리오로 계산됨. 요약 시트는 세 시나리오를 동시에 보여 줌.'
    ws['A2'].font = GREY_F
    ws['A3'], ws['B3'] = '시나리오 선택', '→'
    ws['A3'].font = BOLD
    ws['C3'] = '기본'
    ws['C3'].font, ws['C3'].fill = BLUE_B, SEL_FILL
    dv = DataValidation(type='list', formula1='"보수,기본,낙관"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add('C3')
    ws['D3'] = '(보수·기본·낙관 중 선택)'
    ws['D3'].font = GREY_F
    ws['A4'], ws['C4'] = '선택 번호', '=MATCH(C3,{"보수","기본","낙관"},0)'
    ws['A4'].font = GREY_F
    hdr = ['구분', '항목', '단위', '보수', '기본', '낙관', '적용값(선택)', '근거 성격', '근거·출처', '측정 필요', '팀 초안(비교용, 계산에 안 씀)', '키']
    r = 6
    for i, t in enumerate(hdr):
        c = ws.cell(r, i + 1, t)
        c.font, c.fill = HDR_FONT, HDR_FILL
        c.alignment = Alignment(wrap_text=True, vertical='center')
    r += 1
    for title, rows in SECTIONS:
        c = ws.cell(r, 1, title)
        c.font = BOLD
        for cc in range(1, 13):
            ws.cell(r, cc).fill = SEC_FILL
        r += 1
        for (key, label, unit, fmt, v1, v2, v3, ev, src, meas, draft) in rows:
            model.arow[key] = r
            ws.cell(r, 1, '')
            ws.cell(r, 2, label).alignment = WRAP
            ws.cell(r, 3, unit)
            for j, v in enumerate([v1, v2, v3]):
                c = ws.cell(r, 4 + j, v)
                c.font, c.number_format = BLUE, FMT[fmt]
                c.fill = INPUT_FILL
            c = ws.cell(r, 7, f'=CHOOSE($C$4,D{r},E{r},F{r})')
            c.number_format, c.font, c.fill = FMT[fmt], BOLD, SEL_FILL
            ws.cell(r, 8, ev).alignment = WRAP
            ws.cell(r, 9, src).alignment = WRAP
            c = ws.cell(r, 10, meas)
            if meas:
                c.font = Font(color='C00000', bold=True)
            ws.cell(r, 11, draft).alignment = WRAP
            ws.cell(r, 12, key).font = GREY_F
            for cc in range(1, 13):
                ws.cell(r, cc).border = BOX
            r += 1
    widths = [3, 44, 9, 12, 12, 12, 13, 18, 70, 14, 18, 14]
    for i, w in enumerate(widths):
        ws.column_dimensions[CL(i + 1)].width = w
    ws.freeze_panes = 'D7'

    # ---- 확장 시트 머리(켜기·비교) — 월별 행 위치를 정하기 전에 켜기 행 확정
    wx = wb['확장']
    wx['A1'] = '확장 시나리오 — 단계별로 켜고 끄기(1 = 켬, 0 = 끔)'
    wx['A1'].font = TITLE
    wx['A2'] = ('단계 정의는 「ROOMOVE 확장·경쟁 분석 2026-11-26 v1」의 확장 단계(진입 → 입주 직후 → 거주 중 → 다음 이사)와 고객 확보 통로(제휴)를 따름. '
                '시작 시점·전환율·단가는 가정 시트 G절에서 바꿈.')
    wx['A2'].font = GREY_F
    th = [(1, '단계'), (3, '켜기'), (4, '시작'), (8, '하는 일'), (15, '누가 돈을 내나'), (21, '근거 강도')]
    for col, t in th:
        c = wx.cell(4, col, t)
        c.font, c.fill = HDR_FONT, HDR_FILL
    for cc in range(1, 28):
        wx.cell(4, cc).fill = HDR_FILL
    tr = 5
    start_key = {'e1': 'e1_start', 'e2': 'e2_start', 'e3': 'e3_start', 'e4': 'e4_start'}
    dv2 = DataValidation(type='list', formula1='"1,0"', allow_blank=False)
    wx.add_data_validation(dv2)
    for key, label, val, what, who, ev in TOGGLES:
        model.trow[key] = tr
        wx.cell(tr, 1, label).font = BOLD
        c = wx.cell(tr, 3, val)
        c.font, c.fill = BLUE_B, SEL_FILL
        dv2.add(f'C{tr}')
        wx.cell(tr, 4, f"='가정'!G{model.arow[start_key[key]]}").number_format = '0"개월차"'
        for col, end, txt in [(8, 14, what), (15, 20, who), (21, 27, ev)]:
            wx.cell(tr, col, txt).alignment = WRAP
            wx.merge_cells(start_row=tr, start_column=col, end_row=tr, end_column=end)
        wx.row_dimensions[tr].height = 32
        tr += 1
    wx.cell(tr + 1, 1, ('참고(상위 조사): 미국 Moved는 사람 비서형 B2C로 시작해 임대 운영사 대상 입주 자동화 소프트웨어(B2B2C)로 전환했고 '
                        '2025.11 Paylode 인수까지 영업 중. 전환 이유가 매니저 인건비라는 직접 근거는 없음. '
                        '이 모델에서는 확장3을 켜면 고객은 늘지만 매니저 업무량도 같이 늘어, 자동화율이 낮으면 이익이 커지지 않음.')).alignment = WRAP
    wx.merge_cells(start_row=tr + 1, start_column=1, end_row=tr + 1, end_column=27)
    wx.row_dimensions[tr + 1].height = 34
    cmp_row = tr + 3  # 비교표 시작
    x_month_start = cmp_row + 16

    # ---- 행 배정: 상세(선택)
    starts = {'깔때기': 4, '매출': 27, '비용': 4, '손익': 4, '확장': x_month_start}
    layout(model, '선택', lambda s: s, starts)
    # ---- 행 배정: 시나리오계산(세 블록)
    wc = wb['시나리오계산']
    blk_start = {}
    rr = 4
    for s in SCEN:
        blk_start[s] = rr
        end = layout(model, s, lambda _s: '시나리오계산', {'시나리오계산': rr + 1})
        rr = end['시나리오계산'] + 3

    # ---- 월별 행 쓰기
    write_rows(model, '선택', lambda s: s)
    for s in SCEN:
        c = wc.cell(blk_start[s], 1, f'■ {s} 시나리오 (가정 시트 {SCEN_COL[s]}열 참조)')
        c.font = Font(bold=True, size=12, color='C00000')
        write_rows(model, s, lambda _s: '시나리오계산')
    wc['A1'] = '시나리오계산 — 세 시나리오를 같은 수식으로 동시에 계산(요약 시트용, 직접 고치지 않음)'
    wc['A1'].font = TITLE
    wc['A2'] = '행 구성은 깔때기·매출·비용·손익·확장 시트와 같음. 확장 켜기는 세 시나리오 공통(확장 시트 C열).'
    wc['A2'].font = GREY_F

    for n, title, note in [
        ('깔때기', '고객 깔때기 — 판매량은 방문 × 전환율로만 계산', '선택 시나리오 기준. 전환율 행(들여쓴 행)은 가정 시트 값을 그대로 보여 줌. 관리 고객 = 업체 1곳 이상을 ROOMOVE로 계약한 이사.'),
        ('매출', '매출 — 업종별 거래액 × 수수료율 + 건당 정액', '위: 관리 고객 1건당 단위경제(세 시나리오 + 선택). 아래: 월별 매출(선택 시나리오).'),
        ('비용', '비용 — 이사매니저 블록 · 마케팅 · 변동비 · 고정비', '매니저 수 = 업무량 ÷ (기준 처리 건수 ÷ (1 − 자동화율)), 올림. 매니저 고정비는 계단처럼 늘어남.'),
        ('손익', '손익(월별 36개월) · 손익분기', '손익분기 시점은 C열. 아래 "코어만"은 확장 켜기와 관계없이 확장을 모두 끈 경우.'),
    ]:
        w = wb[n]
        w['A1'], w['A2'] = title, note
        w['A1'].font, w['A2'].font = TITLE, GREY_F
    for n in DETAIL_SHEETS + ['시나리오계산']:
        set_widths(wb[n])
    for n in ['깔때기', '비용', '손익']:
        wb[n].freeze_panes = 'H6'
    wb['시나리오계산'].freeze_panes = 'H4'
    wb['매출'].freeze_panes = None

    # ---- 매출: 1건당 단위경제(세 시나리오 + 선택)
    wr = wb['매출']
    A = lambda k, col: f"'가정'!{col}{model.arow[k]}"
    ue_hdr = ['관리 고객 1건당 단위경제', '단위', '보수', '기본', '낙관', '선택']
    for i, t in enumerate(ue_hdr):
        c = wr.cell(4, i + 1, t)
        c.font, c.fill = HDR_FONT, HDR_FILL
    ue = [
        ('이사 = 이용률 × (거래액 × 수수료율 + 정액)', 'won', lambda c: f'={A("move_use",c)}*({A("move_gmv",c)}*{A("move_take",c)}+{A("move_fixed",c)})'),
        ('입주청소 = 이용률 × 거래액 × 수수료율', 'won', lambda c: f'={A("clean_use",c)}*{A("clean_gmv",c)}*{A("clean_take",c)}'),
        ('부분 시공', 'won', lambda c: f'={A("part_use",c)}*{A("part_gmv",c)}*{A("part_take",c)}'),
        ('큰 공사', 'won', lambda c: f'={A("big_use",c)}*{A("big_gmv",c)}*{A("big_take",c)}'),
        ('인터넷·렌탈 = 가입률 × 건당 정액', 'won', lambda c: f'={A("net_use",c)}*{A("net_fixed",c)}'),
        ('업체 수수료 합계', 'won', None),
        ('(−) 직거래 이탈', 'won', None),
        ('고객 조율 이용료 = 결제 비율 × 금액', 'won', lambda c: f'={A("coord_pay",c)}*{A("coord_fee",c)}'),
        ('1건당 순매출', 'won', None),
        ('(−) 결제·충당·건당 운영비', 'won', None),
        ('(−) 매니저 인건비: 출시 시점 자동화율', 'won', lambda c: f'=-{A("mgr_salary",c)}*(1+{A("mgr_burden",c)})*(1-{A("auto_start",c)})/{A("mgr_cap",c)}'),
        ('(−) 매니저 인건비: 36개월차 자동화율', 'won', lambda c: f'=-{A("mgr_salary",c)}*(1+{A("mgr_burden",c)})*(1-{A("auto_end",c)})/{A("mgr_cap",c)}'),
        ('공헌이익(마케팅 전, 36개월차)', 'won', None),
        ('고객 확보 비용(CAC) = 방문당 광고비 × 유료 비중 ÷ 방문→계약 전환율', 'won',
         lambda c: f'={A("paid_share",c)}*{A("cpc",c)}/({A("conv_signup",c)}*{A("conv_plan",c)}*{A("conv_manage",c)})'),
        ('공헌이익(마케팅 후, 36개월차)', 'won', None),
        ('공헌이익(마케팅 전) ÷ CAC  — 판단 기준 2배 이상', 'dec', None),
        ('매니저 1명 손익분기 처리 건수 = 매니저 월 비용 ÷ (순매출 − 변동비)', 'cnt1', None),
        ('  참고: 매니저 1명 처리 가능 건수(36개월차)', 'cnt1', lambda c: f'={A("mgr_cap",c)}/(1-{A("auto_end",c)})'),
    ]
    base = 5
    for i, (label, fmt, fn) in enumerate(ue):
        r = base + i
        wr.cell(r, 1, label)
        wr.cell(r, 2, '배' if fmt == 'dec' else UNIT.get(fmt, ''))
        for j, col in enumerate(['D', 'E', 'F', 'G']):
            cc = CL(3 + j)
            R = lambda k: f'{cc}{base + k}'
            if fn:
                f = fn(col)
            elif i == 5:
                f = f'=SUM({R(0)}:{R(4)})'
            elif i == 6:
                f = f'=-{R(5)}*{A("leakage",col)}'
            elif i == 8:
                f = f'={R(5)}+{R(6)}+{R(7)}'
            elif i == 9:
                f = f'=-({R(8)}*({A("pg_rate",col)}+{A("reserve",col)})+{A("per_case",col)})'
            elif i == 12:
                f = f'={R(8)}+{R(9)}+{R(11)}'
            elif i == 14:
                f = f'={R(12)}-{R(13)}'
            elif i == 15:
                f = f'=IFERROR({R(12)}/{R(13)},0)'
            elif i == 16:
                f = f'=IF(({R(8)}+{R(9)})>0,{A("mgr_salary",col)}*(1+{A("mgr_burden",col)})/({R(8)}+{R(9)}),"불가")'
            c = wr.cell(r, 3 + j, f)
            c.number_format = FMT[fmt] if fmt != 'dec' else '0.0"배"'
            if j == 3:
                c.fill = SEL_FILL
        if i in (8, 12, 14, 15):
            for cc in range(1, 7):
                wr.cell(r, cc).font = BOLD
    wr.cell(base + len(ue), 1, '매니저 인건비 1건당 = 월 인건비 ÷ 처리 가능 건수(가동률 100% 가정, 실제는 정수 인원이라 더 큼). CAC는 코어 광고 유입 기준.').font = GREY_F
    UE = {'rev': base + 8, 'cm_pre': base + 12, 'cac': base + 13, 'cm_post': base + 14, 'ltv_cac': base + 15, 'be_mgr': base + 16, 'cap36': base + 17}
    assert base + len(ue) + 2 <= starts['매출'], '단위경제 표가 월별 표와 겹침'

    # ---- 확장 시트 비교표(코어만 vs 코어+확장, 세 시나리오)
    def sc(s, key, col):
        sh, r = model.loc[(s, key)]
        return f"'{sh}'!{col}{r}"
    hh = ['코어만 vs 코어 + 켠 확장', '단위', '보수', '기본', '낙관']
    for i, t in enumerate(hh):
        c = wx.cell(cmp_row, i + 1, t)
        c.font, c.fill = HDR_FONT, HDR_FILL
    cmp_items = [
        ('3년 누적 매출: 코어만', 'mw', 'c_rev', 'F'), ('3년 누적 매출: 코어 + 확장', 'mw', 'pl_rev', 'F'),
        ('3년 누적 영업이익: 코어만', 'mw', 'c_op', 'F'), ('3년 누적 영업이익: 코어 + 확장', 'mw', 'pl_op', 'F'),
        ('3년차 영업이익: 코어만', 'mw', 'c_op', 'E'), ('3년차 영업이익: 코어 + 확장', 'mw', 'pl_op', 'E'),
        ('월 손익분기 시점: 코어만', 'mon', 'c_be_month', 'C'), ('월 손익분기 시점: 코어 + 확장', 'mon', 'be_month', 'C'),
        ('3년차 말 필요 매니저: 코어만', 'ppl', 'mgrs_core', 'E'), ('3년차 말 필요 매니저: 코어 + 확장', 'ppl', 'mgrs', 'E'),
        ('3년차 말 정기 청소 고객(확장2)', 'ppl', 'e2_subs', 'E'), ('3년차 말 제휴처(확장3)', 'gok', 'e3_partners', 'E'),
    ]
    for i, (label, fmt, key, col) in enumerate(cmp_items):
        r = cmp_row + 1 + i
        wx.cell(r, 1, label)
        wx.cell(r, 2, UNIT.get(fmt, ''))
        for j, s in enumerate(SCEN):
            c = wx.cell(r, 3 + j, '=' + sc(s, key, col))
            c.number_format = FMT[fmt]
        if '코어 + 확장' in label:
            for cc in range(1, 6):
                wx.cell(r, cc).font = BOLD
    r = cmp_row + 1 + len(cmp_items)
    wx.cell(r, 1, '확장 효과(3년 누적 영업이익 차이)').font = BOLD
    wx.cell(r, 2, '백만원')
    for j in range(3):
        col = CL(3 + j)
        c = wx.cell(r, 3 + j, f'={col}{cmp_row + 4}-{col}{cmp_row + 3}')
        c.number_format, c.font = FMT['mw'], BOLD
    wx.cell(x_month_start - 1, 1, '월별 확장 계산(선택 시나리오)').font = Font(bold=True, size=12)

    # ---- 민감도
    wsn = wb['민감도']
    wsn['A1'] = '민감도 — 36개월차 월 영업이익(코어만, 선택 시나리오): 전환율·수수료율·매니저 처리 건수·자동화율·광고비'
    wsn['A1'].font = TITLE
    wsn['A2'] = ('표 안 값 = 36개월차 한 달의 코어 영업이익(확장 제외). 축 값(파란 글씨)을 바꿔 볼 수 있음. 나머지 가정은 선택 시나리오 값. '
                 '아래 "점검" 칸이 0이면 손익 시트 36개월차와 같은 식임.')
    wsn['A2'].font = GREY_F
    G = lambda k: f"'가정'!$G${model.arow[k]}"
    helpers = [
        ('V36', '36개월차 방문 수', f'=MIN({G("visit_cap")},{G("visit_m1")}*(1+{G("visit_growth")})^35)', 'cnt'),
        ('SP', '방문 → 실행계획 전환율', f'={G("conv_signup")}*{G("conv_plan")}', 'pct'),
        ('RT', '정률 수수료 1건당(이사·청소·시공·공사)', f'={G("move_use")}*{G("move_gmv")}*{G("move_take")}+{G("clean_use")}*{G("clean_gmv")}*{G("clean_take")}+{G("part_use")}*{G("part_gmv")}*{G("part_take")}+{G("big_use")}*{G("big_gmv")}*{G("big_take")}', 'won'),
        ('RF', '정액 1건당(이사 정액·인터넷)', f'={G("move_use")}*{G("move_fixed")}+{G("net_use")}*{G("net_fixed")}', 'won'),
        ('CO', '고객 조율 이용료 1건당', f'={G("coord_pay")}*{G("coord_fee")}', 'won'),
        ('LK', '직거래 이탈', f'={G("leakage")}', 'pct'),
        ('PR', '결제·충당 비율', f'={G("pg_rate")}+{G("reserve")}', 'pct'),
        ('PC', '건당 운영 변동비', f'={G("per_case")}', 'won'),
        ('SF', '매니저 1명 월 비용', f'={G("mgr_salary")}*(1+{G("mgr_burden")})', 'won'),
        ('MN', '최소 매니저 수', f'={G("mgr_min")}', 'cnt'),
        ('MK', '36개월차 유료 광고비', f'=C5*{G("paid_share")}*{G("cpc")}', 'won'),
        ('FX', '36개월차 고정비', f'={G("hc_y3")}*{G("fix_salary")}+{G("server")}+{G("other_fixed")}', 'won'),
        ('CM', '기준: 실행계획 → 업체 계약 전환율', f'={G("conv_manage")}', 'pct'),
        ('CP', '기준: 매니저 처리 건수', f'={G("mgr_cap")}', 'cnt'),
        ('AE', '기준: 36개월차 자동화율', f'={G("auto_end")}', 'pct'),
        ('PL', '회원가입 → 실행계획 전환율', f'={G("conv_plan")}', 'pct'),
        ('PS', '유료 광고 유입 비중', f'={G("paid_share")}', 'pct'),
    ]
    wsn['A4'], wsn['C4'] = '계산 재료(선택 시나리오)', '값'
    wsn['A4'].font = BOLD
    H = {}
    for i, (k, label, f, fmt) in enumerate(helpers):
        r = 5 + i
        H[k] = f'$C${r}'
        wsn.cell(r, 1, label)
        wsn.cell(r, 2, k).font = GREY_F
        wsn.cell(r, 3, f).number_format = FMT[fmt]

    def op_formula(conv, mult, cap, auto, sp=None, mk=None):
        sp = sp or H['SP']
        mk = mk or H['MK']
        cust = f"({H['V36']}*{sp}*{conv})"
        R = f"(({H['RT']}*{mult}+{H['RF']})*(1-{H['LK']})+{H['CO']})"
        mg = f"MAX({H['MN']},ROUNDUP({cust}/({cap}/(1-{auto})),0))"
        return f"={cust}*{R}*(1-{H['PR']})-{cust}*{H['PC']}-{mg}*{H['SF']}-{mk}-{H['FX']}"

    def mgr_formula(conv, cap, auto):
        cust = f"({H['V36']}*{H['SP']}*{conv})"
        return f"=MAX({H['MN']},ROUNDUP({cust}/({cap}/(1-{auto})),0))"

    r0 = 5 + len(helpers) + 1
    chk = r0
    wsn.cell(chk, 1, '점검: 기준값 계산 − 손익 시트 36개월차 코어 영업이익').font = BOLD
    lc = model.loc[('선택', 'c_op')]
    wsn.cell(chk, 3, op_formula(H['CM'], 1, H['CP'], H['AE']) + f"-'{lc[0]}'!{CL(M0 + MONTHS - 1)}{lc[1]}").number_format = '#,##0'

    def table(top, title, rlabel, rvals, rfmt, clabel, cvals, cfmt, cell_fn, out_fmt):
        wsn.cell(top, 1, title).font = Font(bold=True, size=12)
        wsn.cell(top + 1, 1, f'행: {rlabel} / 열: {clabel}').font = GREY_F
        for j, v in enumerate(cvals):
            c = wsn.cell(top + 2, 3 + j, v)
            c.font, c.number_format, c.fill = BLUE_B, cfmt, INPUT_FILL
        for i, v in enumerate(rvals):
            rr_ = top + 3 + i
            c = wsn.cell(rr_, 2, v)
            c.font, c.number_format, c.fill = BLUE_B, rfmt, INPUT_FILL
            for j in range(len(cvals)):
                rc = f'$B{rr_}'
                ccx = f'{CL(3 + j)}${top + 2}'
                c = wsn.cell(rr_, 3 + j, cell_fn(rc, ccx))
                c.number_format = out_fmt
        rng = f'C{top + 3}:{CL(2 + len(cvals))}{top + 2 + len(rvals)}'
        if out_fmt == FMT['mw']:
            wsn.conditional_formatting.add(rng, ColorScaleRule(start_type='min', start_color='F8696B', mid_type='num', mid_value=0, mid_color='FFFFFF', end_type='max', end_color='63BE7B'))
        else:
            wsn.conditional_formatting.add(rng, ColorScaleRule(start_type='min', start_color='FFFFFF', end_type='max', end_color='F4B183'))
        return top + 4 + len(rvals)

    t = r0 + 2
    t = table(t, '표 1. 업체 계약 전환율 × 수수료율 배수 → 36개월차 월 영업이익(백만원)',
              '실행계획 → 업체 계약 전환율', [0.15, 0.20, 0.25, 0.30, 0.35, 0.40, 0.50], '0%',
              '수수료율 배수(이사·청소·시공·공사 정률 수수료율에 곱함, 정액 제외)', [0.5, 0.75, 1.0, 1.25, 1.5, 2.0], '0.00"배"',
              lambda rc, cc: op_formula(rc, cc, H['CP'], H['AE']), FMT['mw'])
    t = table(t, '표 2. 매니저 1명 처리 건수 × 36개월차 자동화율 → 36개월차 월 영업이익(백만원)',
              '매니저 1명 월 처리 건수(자동화 0% 기준)', [8, 12, 15, 20, 25, 30, 40], '0"건"',
              '36개월차 자동화율', [0.0, 0.2, 0.4, 0.5, 0.6, 0.7, 0.8], '0%',
              lambda rc, cc: op_formula(H['CM'], 1, rc, cc), FMT['mw'])
    t = table(t, '표 3. 매니저 1명 처리 건수 × 36개월차 자동화율 → 필요 매니저 수(명, 코어)',
              '매니저 1명 월 처리 건수(자동화 0% 기준)', [8, 12, 15, 20, 25, 30, 40], '0"건"',
              '36개월차 자동화율', [0.0, 0.2, 0.4, 0.5, 0.6, 0.7, 0.8], '0%',
              lambda rc, cc: mgr_formula(H['CM'], rc, cc), '#,##0')
    t = table(t, '표 4. 방문 → 회원가입 전환율 × 유료 방문 1회당 광고비 → 36개월차 월 영업이익(백만원)',
              '방문 → 회원가입 전환율', [0.04, 0.06, 0.08, 0.10, 0.12, 0.15, 0.20], '0%',
              '유료 방문 1회당 광고비(원)', [300, 500, 1000, 1500, 2000, 3000], '#,##0"원"',
              lambda rc, cc: op_formula(H['CM'], 1, H['CP'], H['AE'], sp=f"({rc}*{H['PL']})", mk=f"({H['V36']}*{H['PS']}*{cc})"), FMT['mw'])
    wsn.column_dimensions['A'].width = 44
    wsn.column_dimensions['B'].width = 10
    for j in range(8):
        wsn.column_dimensions[CL(3 + j)].width = 12

    # ---- 요약
    wsu = wb['요약']
    wsu['A1'] = 'ROOMOVE 재무 시뮬레이션 v1 — 요약'
    wsu['A1'].font = TITLE
    wsu['A2'] = ('세 시나리오를 동시에 계산한 결과(코어 + 확장 시트에서 켠 단계). 모든 전환율·처리 건수는 팀 추정이며 측정 전 크기 가늠용. '
                 '"선택" 열은 상세 시트 값(가정 시트에서 고른 시나리오).')
    wsu['A2'].font = GREY_F
    wsu['A3'] = '현재 선택 시나리오'
    wsu['C3'] = "='가정'!C3"
    wsu['C3'].font, wsu['C3'].fill = BOLD, SEL_FILL
    sh_ = ['지표', '단위', '보수', '기본', '낙관', '선택']
    for i, t_ in enumerate(sh_):
        c = wsu.cell(5, i + 1, t_)
        c.font, c.fill = HDR_FONT, HDR_FILL

    def det(key, col):
        sh, r = model.loc[('선택', key)]
        return f"'{sh}'!{col}{r}"
    items = [
        ('매출·이익', None),
        ('1년차 매출', 'mw', 'pl_rev', 'C'), ('2년차 매출', 'mw', 'pl_rev', 'D'), ('3년차 매출', 'mw', 'pl_rev', 'E'),
        ('3년 누적 매출', 'mw', 'pl_rev', 'F'), ('3년 누적 거래액(GMV)', 'mw', 'gmv_total', 'F'),
        ('1년차 영업이익', 'mw', 'pl_op', 'C'), ('2년차 영업이익', 'mw', 'pl_op', 'D'), ('3년차 영업이익', 'mw', 'pl_op', 'E'),
        ('3년 누적 영업이익', 'mw', 'pl_op', 'F'),
        ('손익분기', None),
        ('월 손익분기 시점', 'mon', 'be_month', 'C'), ('누적 손익 회복 시점', 'mon', 'be_cum', 'C'),
        ('36개월차 손익분기 필요 관리 고객(월)', 'cnt', 'be_cust', 'E'), ('36개월차 손익분기 필요 방문(월)', 'vis', 'be_visits', 'E'),
        ('36개월차 실제 관리 고객(월)', 'cnt', 'cust_total', 'AQ'), ('36개월차 실제 방문(월)', 'vis', 'visits', 'AQ'),
        ('이사매니저(핵심 리스크)', None),
        ('3년차 말 필요 매니저 수', 'ppl', 'mgrs', 'E'), ('3년차 말 매니저 1명 처리 건수', 'cnt1', 'mgr_util', 'E'),
        ('3년차 매니저 인건비 ÷ 매출', 'pct', None, None),
        ('단위경제', None),
        ('관리 고객 1건당 순매출', 'won', 'UE', 'rev'), ('1건당 공헌이익(마케팅 전, 36개월차 자동화율)', 'won', 'UE', 'cm_pre'),
        ('고객 확보 비용(CAC, 유료 광고 기준)', 'won', 'UE', 'cac'), ('1건당 공헌이익(마케팅 후)', 'won', 'UE', 'cm_post'),
        ('공헌이익(마케팅 전) ÷ CAC (기준 2배 이상)', 'x', 'UE', 'ltv_cac'),
        ('매니저 1명 손익분기 처리 건수', 'cnt1', 'UE', 'be_mgr'), ('매니저 1명 처리 가능 건수(36개월차)', 'cnt1', 'UE', 'cap36'),
        ('3년차 말 수도권 대상 가구 점유율(연환산)', 'pct', 'share', 'E'),
        ('코어만(확장 모두 끔)', None),
        ('3년 누적 매출: 코어만', 'mw', 'c_rev', 'F'), ('3년 누적 영업이익: 코어만', 'mw', 'c_op', 'F'),
        ('월 손익분기 시점: 코어만', 'mon', 'c_be_month', 'C'), ('3년차 말 필요 매니저: 코어만', 'ppl', 'mgrs_core', 'E'),
    ]
    r = 6
    rowmap = {}
    for it in items:
        if it[1] is None:
            c = wsu.cell(r, 1, it[0])
            c.font = BOLD
            for cc in range(1, 7):
                wsu.cell(r, cc).fill = SEC_FILL
            r += 1
            continue
        label, fmt, key, col = it
        rowmap[label] = r
        wsu.cell(r, 1, label)
        wsu.cell(r, 2, '배' if fmt == 'x' else UNIT.get(fmt, ''))
        for j, s in enumerate(SCEN + ['선택']):
            if key is None:  # 매니저 인건비 ÷ 매출
                if s == '선택':
                    f = f'=IFERROR({det("mgr_cost", "E")}/{det("pl_rev", "E")},0)'
                else:
                    f = f'=IFERROR({sc(s, "mgr_cost", "E")}/{sc(s, "pl_rev", "E")},0)'
            elif key == 'UE':
                f = f"='매출'!{CL(3 + j)}{UE[col]}"
            elif s == '선택':
                f = '=' + det(key, col)
            else:
                f = '=' + sc(s, key, col)
            c = wsu.cell(r, 3 + j, f)
            c.number_format = '0.0"배"' if fmt == 'x' else FMT[fmt]
            if s == '선택':
                c.fill = SEL_FILL
        if label in ('3년 누적 매출', '3년 누적 영업이익', '월 손익분기 시점', '3년차 말 필요 매니저 수'):
            for cc in range(1, 7):
                wsu.cell(r, cc).font = BOLD
        r += 1
    wsu.cell(r, 1, '점검: 선택 열 = 같은 시나리오 열(3년 누적 매출)').font = GREY_F
    rr_ = rowmap['3년 누적 매출']
    wsu.cell(r, 6, f'=IF(ABS(F{rr_}-INDEX(C{rr_}:E{rr_},\'가정\'!C4))<1,"일치","불일치")').font = GREY_F
    wsu.cell(r + 1, 1, '단위경제 행은 매출 시트 위쪽 표(세 시나리오)를 그대로 가져옴. 매니저 인건비는 가동률 100% 가정이라 실제(정수 인원)보다 작게 나옴.').font = GREY_F
    for col, w in zip('ABCDEF', [40, 8, 12, 12, 12, 12]):
        wsu.column_dimensions[col].width = w

    # 차트 1: 시나리오별 연도별 영업이익
    ch = BarChart()
    ch.type = 'col'
    ch.title = '시나리오별 연도별 영업이익(코어 + 켠 확장)'
    ch.y_axis.title = '백만원'
    ch.y_axis.numFmt = '#,##0,,'
    y1, y3 = rowmap['1년차 영업이익'], rowmap['3년차 영업이익']
    for j, s in enumerate(SCEN):
        se = Series(Reference(wsu, min_col=3 + j, min_row=y1, max_row=y3), title=s)
        ch.series.append(se)
    ch.set_categories(Reference(wsu, min_col=1, min_row=y1, max_row=y3))
    ch.height, ch.width = 7.5, 15
    wsu.add_chart(ch, 'H5')
    # 차트 2: 시나리오별 필요 매니저 수(월별)
    lc2 = LineChart()
    lc2.title = '시나리오별 필요 이사매니저 수(월별)'
    lc2.y_axis.title = '명'
    lc2.x_axis.title = '개월차'
    for s in SCEN:
        shn, rrr = model.loc[(s, 'mgrs')]
        se = Series(Reference(wc, min_col=M0, max_col=M0 + MONTHS - 1, min_row=rrr), title=s)
        lc2.series.append(se)
    hr = model.hdr[(SCEN[0], '시나리오계산')] + 1
    lc2.set_categories(Reference(wc, min_col=M0, max_col=M0 + MONTHS - 1, min_row=hr))
    lc2.height, lc2.width = 7.5, 15
    wsu.add_chart(lc2, 'H21')

    # ---- 안내
    wa = wb['안내']
    lines = [
        ('ROOMOVE 재무 시뮬레이션 v1 (2026-11-26)', TITLE),
        ('아이템: ROOMOVE(이사 실행계획·이사매니저). 수익 = 계약 성사 시 업체 수수료(시공·청소 정률, 이사는 법 문제로 시나리오별 처리) + 인터넷·렌탈 건당 정액 + 확장 매출', None),
        ('', None),
        ('쓰는 법', BOLD),
        ('1. 가정 시트 C3에서 시나리오(보수·기본·낙관)를 고르면 깔때기·매출·비용·손익·확장 시트가 그 시나리오로 계산됨', None),
        ('2. 요약 시트는 세 시나리오를 동시에 보여 줌(시나리오계산 시트가 같은 수식으로 세 번 계산)', None),
        ('3. 확장 시트 C열에서 확장 단계를 1(켬)/0(끔)으로 바꿈. 요약·확장 시트에 "코어만"과 "코어 + 켠 확장"이 나란히 나옴', None),
        ('4. 민감도 시트의 축 값(파란 글씨)을 바꿔 전환율·수수료율·매니저 처리 건수·자동화율의 영향을 봄', None),
        ('', None),
        ('색 규칙', BOLD),
        ('파란 글씨 = 입력값(가정 시트 D~F열, 확장 시트 켜기, 민감도 축). 검은 글씨 = 수식(고치지 않음). 노란 칸 = 선택 시나리오 값', None),
        ('금액: 월별 표는 원 단위로 계산하고 백만원으로 표시(예: 12.3 = 1,230만원). 단위경제 표는 원 단위', None),
        ('', None),
        ('시트 구성', BOLD),
        ('요약: 시나리오별 3년 매출·영업이익·손익분기 시점·필요 매니저 수, 차트 2개', None),
        ('가정: 입력값 전부. 항목마다 단위·근거 성격(공식 자료/공시/언론 보도/업체 자료/학술 자료/팀 추정)·출처·측정 필요·팀 초안 값(비교용)', None),
        ('깔때기: 방문 → 회원가입 → 실행계획 → 업체 계약(관리 고객) → 업종별 계약. 판매량은 방문 × 전환율로만 계산', None),
        ('매출: 1건당 단위경제(세 시나리오) + 업종별 거래액 × 수수료율 + 건당 정액, 직거래 이탈 차감', None),
        ('비용: 이사매니저 블록(처리 건수·자동화율 → 필요 인원 → 인건비), 마케팅, 변동비, 고정비', None),
        ('손익: 월별 36개월(연도 합계 포함), 손익분기 시점, 손익분기 필요 관리 고객·방문 수, 코어만 비교', None),
        ('확장: 확장1 입주 직후 연계 판매 / 확장2 거주 중 정기 청소 / 확장3 제휴 채널(B2B2C) / 확장4 다음 이사 재이용', None),
        ('민감도: 36개월차 월 영업이익(코어) 2변수 표 2개 + 필요 매니저 수 표 1개', None),
        ('시나리오계산: 요약용 계산 시트(직접 고치지 않음)', None),
        ('', None),
        ('팀 초안과 다른 점(비교용으로만 둠)', BOLD),
        ('이사 10% 성과 수수료: 운송주선 허가 문제(2026.8 개정법)로 보수·기본에서는 빼고, 기본은 허가 업체 제휴 정액(건당 6만원), 낙관만 허가 취득 후 10%', None),
        ('입주청소 10%·시공 3%: 기본 시나리오에 그대로 쓰고 보수·낙관에서 위아래로 움직임', None),
        ('인터넷·렌탈 건당 정액: 초안에 금액이 없어 아정당 공시 기반 추정(기본 8만원)으로 둠', None),
        ('', None),
        ('주의', BOLD),
        ('방문·전환율·매니저 처리 건수·자동화율·시공 단가·이용률은 근거가 없거나 약한 팀 추정이며, 가정 시트 J열 "측정 필요" 항목이 다음 90일 측정 대상', None),
        ('관리 고객 수는 기댓값이라 소수로 나옴. 매출 인식 시점(계약 ↔ 이사 완료의 시차)과 계절성은 v1에서 반영하지 않음', None),
        ('빌더: 20-ROOMOVE/tools/build_finance_model.py (행 추가·수정은 ROWS 목록에서, 시나리오계산 시트도 같이 바뀜)', None),
    ]
    for i, (t_, f_) in enumerate(lines):
        c = wa.cell(1 + i, 1, t_)
        if f_:
            c.font = f_
    wa.column_dimensions['A'].width = 150

    for w in wb.worksheets:
        w.sheet_view.zoomScale = 90
    wb['요약'].sheet_properties.tabColor = 'C00000'
    wb['가정'].sheet_properties.tabColor = '0000FF'
    wb['시나리오계산'].sheet_properties.tabColor = '808080'
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    return model


if __name__ == '__main__':
    build(sys.argv[1])
    print('saved', sys.argv[1], '-> 계산값 저장: python 20-ROOMOVE/tools/recalc_excel.py <같은 경로>')
