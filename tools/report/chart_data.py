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
    # 5. 침투 시나리오
    'c5_scenario.png': dict(
        kind='vbar',
        labels=['침투 0.5%\n(약 2.03만 명)', '침투 1%\n(약 4.06만 명)', '침투 3%\n(약 12.2만 명)'],
        values=[24, 49, 146], unit='억 원', vmax=175,
        note='심한 월경통 응답 여성 약 406만 명(추정) × 누적 침투율 × 본체 12만 원 — 규모 예시, 연간 매출 아님'),
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
        values=[2.73, 3.80, 4.18, 4.05, 4.88, 5.65, 4.57, 3.70, 3.95, 3.49], unit='', vmax=6.5,
        note='마사지용 기기(HSK 9019102000) 수입액, 단위 억 달러'),
}
