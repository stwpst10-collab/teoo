"""Chart data shared by charts.py (matplotlib PNG for docx) and build_html.py (inline SVG).

Only numbers and labels live here. Each figure's source line and grade stay in report_content.py.
"""

CHARTS = {
    # 1. 심한 월경 증상 성인의 대처 (질병관리청 2022)
    'c1_coping.png': dict(
        kind='hbar',
        labels=['약국 진통제', '아무것도 안 함', '건강기능식품·민간요법', '한의원', '산부인과·비뇨의학과 진료'],
        values=[59.2, 23.5, 18.2, 4.8, 28.5], unit='%', vmax=75,
        note='심한 월경 증상(VAS 7점 이상)이 1개 이상인 성인 1,266명, 단위 %'),
    # 2. 심한 증상 유병률
    'c2_prevalence.png': dict(
        kind='hbar',
        labels=['성인 심한 월경전 증후군', '성인 심한 월경통', '청소년 심한 월경통', '청소년 심한 월경전 증후군'],
        values=[46.1, 41.6, 42.6, 29.3], unit='%', vmax=60,
        note='월경 중인 청소년 961명·성인 2,137명, VAS 7점 이상 비율, 단위 %'),
    # 3. 펨테크 시장 기관별 2024 추정
    'c3_femtech.png': dict(
        kind='vbar',
        labels=['Fortune Business\nInsights', 'Grand View\nResearch', 'Arizton\n(2023)', 'Global Market\nInsights'],
        values=[78.1, 392.9, 518.0, 602.0], unit='억$', vmax=700,
        note='글로벌 펨테크 시장 규모 추정(2024), 단위 억 달러 — 기관별 정의 상이'),
    # 4. 3개년 계획: 매출과 세전 손익 (같은 단위, 하나의 축)
    'c4_plan.png': dict(
        kind='group',
        labels=['1년차', '2년차', '3년차'],
        series=[('매출', [0.70, 3.00, 8.00], ['0.70', '3.00', '8.00']),
                ('세전 단순 손익', [-0.965, -0.44, 0.75], ['-0.965', '-0.44', '+0.75'])],
        vmin=-1.6, vmax=9,
        note='린캔버스 계획 가정(실적 아님), 단위 억 원, 부가세 제외'),
    # 5. 첫해 판매 규모 시나리오 (깔때기: 방문 × 전환율 × 가격)
    'c5_scenario.png': dict(
        kind='vbar',
        labels=['보수\n(구매자 140명)', '기본\n(구매자 960명)', '낙관\n(구매자 4,440명)'],
        values=[0.12, 1.03, 5.73], unit='억 원', vmax=7, decimals=2,
        note='첫해 매출(소비자가, 억 원) = 방문(미측정) × 구매 전환율 × 가격 — 계산값(D), 예측 아님'),
    # 6. 대형 홈헬스케어 2사 매출 추이 (같은 단위, 하나의 축)
    'c6_homehealth.png': dict(
        kind='group',
        labels=['2021', '2022', '2023', '2024', '2025'],
        series=[('세라젬', [6671, 7502, 5847, 5460, 5498], None),
                ('바디프랜드', [5913, 5220, 4197, 4369, 4226], None)],
        vmin=0, vmax=8800,
        note='매출 추이, 단위 억 원'),
    # 7. 국내 홈뷰티 디바이스 시장
    'c7_beauty.png': dict(
        kind='vbar', labels=['2018', '2022', '2030(전망)'], values=[0.5, 1.6, 3.4], unit='조 원', vmax=4.0,
        note='국내 홈뷰티 디바이스 시장 규모, 단위 조 원'),
    # 8. 마사지용 기기(HSK 9019102000) 수입액 추이 — 관세청
    'c8_imports.png': dict(
        kind='vbar', labels=['2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'],
        values=[2.731, 3.802, 4.175, 4.046, 4.877, 5.647, 4.574, 3.696, 3.946, 3.489], unit='', vmax=6.5,
        note='마사지용 기기(HSK 9019102000) 수입액, 단위 억 달러'),
    # 9. 확장 지도 (1장 1.3, 2장 2.5): 숫자 없음, 단계·조건만. evidence: True = 기존 근거, False = 조사 중
    'c9_expansion.png': dict(
        kind='map', note='무대: 소형 디바이스(착용·휴대형). 칸 안에는 시장 규모를 넣지 않음(합산 금지)',
        rows=['고객 상황', '제품(예)', '같이 가져가는 것', '넘어갈 조건'],
        stages=[
            dict(key='①', title='생리 지연 고민', evidence=True,
                 cells=['생리가 늦어질까|신경 쓰이는 날', '발목 지압 밴드|(류, 수동)', '시작: 고객 원기록,|첫 판매 채널', '첫 고객 문제 확인,|재사용·유료 전환']),
            dict(key='②', title='생리 전후 케어', evidence=True,
                 cells=['생리 전후 배·허리가|불편한 날', '발목 + 온열,|복부·허리(온)', '같은 고객, 발목|위치 지식, 채널', '생리 전후 반복 사용,|기존 제품 불만 확인']),
            dict(key='③', title='바디 웰니스|(붓기·무거움)', evidence=False,
                 cells=['다리가 무겁고 부은|느낌이 신경 쓰이는 날', '종아리·하체 착용', '착용·온열·압박|부품, 사용 기록', '주기와 무관한|일상 사용 확인']),
            dict(key='④-a', title='홈헬스', evidence=True,
                 cells=['목·어깨·허리가 뭉친 날', '목·어깨·휴대(결·틈)', '부품·인증 표현, 제조 맥락', '재구매·추천, 부위 확장 반응']),
            dict(key='④-b', title='홈뷰티(바디 먼저)', evidence=True,
                 cells=['바디라인·붓기가 신경 쓰이는 날', '바디 착용형(선)', '올리브영 채널, 같은 고객', '뷰티 고객의 구매 의향']),
            dict(key='④-c', title='웨어러블·IT(개인화 후)', evidence=False,
                 cells=['내 주기·컨디션에 맞춰 쓰고 싶을 때', '기록·추천(맥·숨)', '사용·주기 기록(동의 전제)', '데이터 동의율, 기록 → 재구매']),
        ],
        gate='조건 미충족 → 고침·바꿈·멈춤',
        contrast='대형 디바이스(안마의자·척추 의료기기): 비교 대상일 뿐, 우리 길 아님'),
}
