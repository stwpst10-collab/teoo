// ROOMOVE 확장·경쟁 분석 발표 자료 + 확장 논리 1장 생성
const pptxgen = require('pptxgenjs');
const path = require('path');
const OUT = process.argv[2];

const F = 'Malgun Gothic';
const NAVY = '15233B', BLUE = '2A78D6', ORANGE = 'EB6834', INK = '1F2937', MUTED = '6B7280',
  LINE = 'D7DEE8', TINT = 'EEF4FC', TINT2 = 'FDF0EA', WHITE = 'FFFFFF', GREEN = '2E7D5B';

function setup(pres) {
  pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
  pres.theme = { headFontFace: F, bodyFontFace: F };
  pres.author = '사내벤처';
  pres.defineSlideMaster({
    title: 'CONTENT', background: { color: WHITE },
    objects: [
      { placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.35, w: 12.1, h: 0.75, fontFace: F, fontSize: 22, bold: true, color: NAVY, align: 'left', valign: 'top', margin: 0 }, text: '' } },
      { text: { text: 'ROOMOVE 확장·경쟁 분석', options: { x: 0.6, y: 7.08, w: 5, h: 0.25, fontFace: F, fontSize: 9, color: MUTED, margin: 0 } } },
    ],
    slideNumber: { x: 12.3, y: 7.08, w: 0.5, h: 0.25, fontFace: F, fontSize: 9, color: MUTED, align: 'right' },
  });
  pres.defineSlideMaster({ title: 'DARK', background: { color: NAVY }, objects: [] });
}

const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, fontFace: F, color: INK, margin: 0, valign: 'top' }, o));
function title(s, text, sub) {
  s.addText(text, { placeholder: 'title' });
  if (sub) T(s, sub, { x: 0.6, y: 1.05, w: 12.1, h: 0.4, fontSize: 14, color: MUTED });
}
function source(s, text) { T(s, '자료: ' + text, { x: 0.6, y: 6.72, w: 12.1, h: 0.32, fontSize: 9, color: MUTED, valign: 'bottom' }); }
function box(s, x, y, w, h, fill, line) {
  s.addShape('roundRect', { x, y, w, h, fill: { color: fill }, line: { color: line || fill, width: 1 }, rectRadius: 0.08 });
}
function tag(s, x, y, text, color) {
  s.addShape('roundRect', { x, y, w: 1.55, h: 0.3, fill: { color }, line: { color }, rectRadius: 0.15 });
  T(s, text, { x, y, w: 1.55, h: 0.3, fontSize: 10, color: WHITE, bold: true, align: 'center', valign: 'middle' });
}
function table(s, rows, opt) {
  const head = rows[0].map(t => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY } } }));
  const body = rows.slice(1).map((r, i) => r.map((t, j) => ({ text: t, options: Object.assign({ fill: { color: i % 2 ? 'F7F9FC' : WHITE } }, (opt.hl && opt.hl(i, j)) || {}) })));
  s.addTable([head, ...body], Object.assign({ fontFace: F, fontSize: 11, color: INK, border: { type: 'solid', color: LINE, pt: 0.75 }, valign: 'middle', margin: [3, 6, 3, 6] }, opt.t));
}
function stat(s, x, y, w, big, label, color) {
  T(s, big, { x, y, w, h: 0.75, fontSize: 30, bold: true, color: color || BLUE, valign: 'bottom' });
  T(s, label, { x, y: y + 0.8, w, h: 0.75, fontSize: 12, color: INK });
}
function arrowR(s, x, y, w) { s.addShape('line', { x, y, w, h: 0, line: { color: MUTED, width: 1.5, endArrowType: 'triangle' } }); }

// ───────────────── 1장: 확장 논리 ─────────────────
function logicSlide(pres, s) {
  title(s, '이사 회사가 아니라, 이사를 시작으로 주거 생애주기를 잇는 플랫폼', '확장 논리(검토안): 이사로 만나 → 입주 직후·거주 중 서비스로 넓히고 → 쌓인 집 이력으로 보장·IoT, 다음 주거 이벤트까지');
  // 대전제
  box(s, 0.6, 1.55, 12.1, 0.62, TINT, TINT);
  T(s, [
    { text: '대전제  ', options: { bold: true, color: BLUE } },
    { text: '주거 이벤트에서 고객을 먼저 만나고 그 뒤 집 관련 지출을 잇는 플랫폼은 거래가 줄어도 매출을 지킴. ' },
    { text: '贝壳 2025 인테리어·임대 비중 약 39%, 중개 감소에도 총매출 +1.2%(공시)', options: { color: MUTED } },
  ], { x: 0.8, y: 1.6, w: 11.8, h: 0.52, fontSize: 13, valign: 'middle' });
  // 5단계
  const st = [
    ['1', '이사·입주 일정', '현재', '업종 사이 일정 조율, 계약 뒤 등기·전입 기한 안내', '시공·청소 업체 수수료', '문제 근거 중간'],
    ['2', '입주 직후 홈서비스', '', '청소·설치·인터넷·수리를 이사 일정 안에서 순서대로', '업체 성사 수수료', '추가 이용 측정 전'],
    ['3', '거주 중 반복·자산 기록', '반복 매출', '가전·설비 목록으로 정기 관리·교체·소모품 제안, 이력 축적', '정기 이용료, 업체·소모품 제휴 수수료', '전환율 측정 전'],
    ['4', '주거 보장·IoT', '장기', '쌓인 이력으로 보험사와 주택 보증, IoT 연동', '보험사 판매 수수료', '해외 공시(Porch·Frontdoor)'],
    ['5', '다음 주거 이벤트', '제휴사 공급', '재이사·임대·매매 지원, 중개사무소·건설사에 일정 기능', '소프트웨어·제휴 수수료', '해외 사례(Updater)'],
  ];
  const x0 = 0.6, w = 2.3, g = 0.15, y = 2.4, h = 2.35;
  st.forEach((r, i) => {
    const x = x0 + i * (w + g);
    const last = i === 3;
    box(s, x, y, w, h, last ? TINT2 : WHITE, last ? ORANGE : LINE);
    s.addShape('ellipse', { x: x + 0.15, y: y + 0.15, w: 0.42, h: 0.42, fill: { color: last ? ORANGE : (i === 0 ? NAVY : BLUE) }, line: { color: WHITE } });
    T(s, r[0], { x: x + 0.15, y: y + 0.15, w: 0.42, h: 0.42, fontSize: 14, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    T(s, r[1], { x: x + 0.65, y: y + 0.12, w: w - 0.75, h: 0.5, fontSize: 13, bold: true, color: NAVY, valign: 'middle' });
    if (r[2]) T(s, r[2], { x: x + 0.65, y: y + 0.55, w: w - 0.75, h: 0.22, fontSize: 9.5, bold: true, color: last ? ORANGE : BLUE });
    T(s, r[3], { x: x + 0.15, y: y + 0.85, w: w - 0.3, h: 0.7, fontSize: 11 });
    T(s, [{ text: '수익원  ', options: { bold: true, color: MUTED } }, { text: r[4] }], { x: x + 0.15, y: y + 1.55, w: w - 0.3, h: 0.4, fontSize: 10 });
    T(s, [{ text: '근거  ', options: { bold: true, color: MUTED } }, { text: r[5] }], { x: x + 0.15, y: y + 1.92, w: w - 0.3, h: 0.35, fontSize: 10 });
  });
  // 플라이휠
  T(s, '플라이휠: 고객이 늘수록 기록이 쌓이고, 기록이 쌓일수록 제안·보장이 정확해짐', { x: 0.6, y: 4.95, w: 7.4, h: 0.3, fontSize: 12, bold: true, color: NAVY });
  const fw = ['이사 고객 증가', '집 단위 이력 축적\n(가전·설비·서비스 이력)','제안·보장 가격 정확도 상승', '반복 매출↑ 확보 비용↓'];
  fw.forEach((t, i) => {
    const x = 0.6 + i * 1.87;
    box(s, x, 5.35, 1.6, 0.95, i === 1 ? NAVY : TINT, i === 1 ? NAVY : TINT);
    T(s, t, { x: x + 0.07, y: 5.35, w: 1.46, h: 0.95, fontSize: 10.5, bold: true, align: 'center', valign: 'middle', color: i === 1 ? WHITE : NAVY });
    if (i < 3) arrowR(s, x + 1.62, 5.82, 0.23);
  });
  T(s, '↺ 다시 고객 증가', { x: 0.6, y: 6.35, w: 7.2, h: 0.28, fontSize: 10, color: MUTED, align: 'right' });
  // 성립 조건
  box(s, 8.2, 4.95, 4.5, 1.7, 'F7F9FC', LINE);
  T(s, '성립 조건과 먼저 측정할 것', { x: 8.38, y: 5.03, w: 4.2, h: 0.3, fontSize: 12, bold: true, color: ORANGE });
  T(s, [
    { text: '시공·재고·보험 위험은 직접 지지 않음(업체·보험사가 부담)', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '고객 확보 비용 7만원↓, 추가 이용 40%↑, 재구매 25%↑', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '정기 서비스 전환율(12개월 재이용 30% 이상)', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '입주 시 자산 정보 등록률, 보장 안내 수락률', options: { bullet: { indent: 14 } } },
  ], { x: 8.38, y: 5.36, w: 4.2, h: 1.25, fontSize: 10.5, paraSpaceAfter: 2 });
  source(s, 'KE Holdings 2025 연간실적, Porch 10-K(FY2025), Frontdoor 10-K·연차보고서(2025), Updater 공시, 청소연구소 발표, 당사 정리. 모든 단계의 함께 쓰는 비율·전환율은 측정 전 가설');
}

// ───────────────── 전체 발표 자료 ─────────────────
function fullDeck(pres) {
  // 1. 표지
  let s = pres.addSlide({ masterName: 'DARK' });
  T(s, '사내벤처 분석  |  2026.11', { x: 0.8, y: 1.6, w: 8, h: 0.4, fontSize: 14, color: 'AFC6E9' });
  T(s, 'ROOMOVE 확장·경쟁 분석', { x: 0.8, y: 2.15, w: 11.5, h: 1.0, fontSize: 40, bold: true, color: WHITE });
  T(s, '이사를 입구로 주거 이동의 실행 단계를 모두 잇는 플랫폼', { x: 0.8, y: 3.25, w: 11.5, h: 0.6, fontSize: 20, color: WHITE });
  T(s, '계약 뒤 대출·등기·전입·이사·시공·배송은 잔금·입주일 전후에 몰리지만, 이를 하나로 잇는 곳은 공공·민간 모두 공개 자료로 확인되지 않음', { x: 0.8, y: 4.4, w: 10.5, h: 0.8, fontSize: 14, color: 'C9D6EA' });
  T(s, '본 자료의 생애가치·수수료는 당사 추정, 근거 성격은 각 쪽 자료 줄에 표기', { x: 0.8, y: 6.6, w: 10, h: 0.3, fontSize: 10, color: '8FA3C2' });

  // 2. 핵심 요약
  s = pres.addSlide({ masterName: 'CONTENT' });
  title(s, '이사를 입구로 계약 이후 절차와 입주 후 지출을 잇는 주거 이동 플랫폼', '투자포인트 요약');
  [['351.8만건', '2025 주택 매매 72.6만 + 전월세 279.2만건\n(갱신 포함, 이사 가구 수와 다름)'],
   ['약 39%', '중국 贝壳 2025 인테리어·임대 매출 비중\n중개 감소에도 총매출 +1.2%'],
   ['약 2배', '미국 기존 주택 구매 이사 가구의 첫해\n가구·가전·수리 지출(비이사 자가 대비)'],
   ['98만~113만원', '고객 1명 5년 누적 수수료(거주 중 서비스 이용 시)\n안 쓰면 22만~33만원, 당사 추정']].forEach((r, i) => {
    stat(s, 0.6 + i * 3.05, 1.6, 2.85, r[0], r[1], i === 3 ? ORANGE : BLUE);
  });
  const pts = [
    ['이사 회사가 아니라 주거 생애주기 플랫폼', '이사를 입구로 고객과 집·가전·설비를 함께 확보하고 서비스 이력을 시간 순으로 쌓음. 요청에 응답하지 않고 필요한 서비스를 먼저 제안하는 것이 목표(가설)'],
    ['이사는 실행이 한 날짜로 몰리는 지점', '대출 실행·등기·전입·이사·시공·배송이 잔금·입주일 전후에 맞물리나 끝까지 잇는 곳은 공개 자료로 확인 못 함'],
    ['오늘의집과 같은 목표, 반대 방향', '오늘의집은 커머스 고객을 시공·이사로, ROOMOVE는 이사 고객을 주거 생활로. 구매처·업체를 묻지 않는 중립적 일정 관리'],
    ['누적 매출은 반복 서비스에서', '정기 청소·가전 관리가 붙으면 5년 수수료 약 3.5~4.4배. 해외 확장 성공 사례도 모두 매년 갱신 상품 보유'],
  ];
  pts.forEach((p, i) => {
    const y = 3.4 + i * 0.8;
    s.addShape('ellipse', { x: 0.6, y: y + 0.05, w: 0.32, h: 0.32, fill: { color: BLUE }, line: { color: BLUE } });
    T(s, String(i + 1), { x: 0.6, y: y + 0.05, w: 0.32, h: 0.32, fontSize: 11, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    T(s, p[0], { x: 1.1, y, w: 3.9, h: 0.6, fontSize: 14, bold: true, color: NAVY, valign: 'middle' });
    T(s, p[1], { x: 5.0, y, w: 7.7, h: 0.6, fontSize: 12, valign: 'middle' });
  });
  source(s, '국토교통부 주택통계(2026.1.30), KE Holdings 2025 연간실적(HKEX), NAHB(2026.8), 당사 추정');

  // 3. 경쟁 구도
  s = pres.addSlide({ masterName: 'CONTENT' });
  title(s, '견적 연결은 이미 흔하고, 업종 사이 일정 조율은 맡는 곳이 없음', '경쟁 구도: 기능별 담당 범위');
  table(s, [
    ['기능', '하는 곳', '수준'],
    ['여러 업종 견적·신청을 한곳에서', '짐싸, 이사타임, 오늘의집, 헬로입주, 홈닉, 아파트아이, 아정당', '이미 흔함'],
    ['도면 위 가구 배치', '오늘의집 3D 인테리어, 아키스케치, 한샘 상담·홈닉 주택형별 제안', '도구 존재(기본 무료 포함)'],
    ['한 회사 안 여러 공정 일정 관리', '삼성 인테리어핏, IKEA 주방, 오늘의집 직접 시공', '자사 상품·자사 시공 한정'],
    ['같은 업종 안 업체 재배정', '아정당', '업종 안 한정'],
    ['업종 사이 순서표·업체별 요청서·지연 시 일정 재조정', '공개 자료로 확인된 곳 없음', 'ROOMOVE 목표 영역'],
  ], { t: { x: 0.6, y: 1.6, w: 12.1, colW: [3.6, 5.9, 2.6], fontSize: 12, rowH: 0.48 },
    hl: (i) => i === 4 ? { bold: true, color: ORANGE, fill: { color: TINT2 } } : null });
  box(s, 0.6, 4.75, 12.1, 1.6, TINT, TINT);
  T(s, [
    { text: '시사점  ', options: { bold: true, color: BLUE } },
    { text: '가구 배치 기능만으로는 차별화 어려움(어반베이스 2024.7 폐업, 집꾸미기 2025 종료). 차별화 후보는 실측 치수로 "가져갈지" 판단한 결과가 이사 품목표·사다리차·시공 치수·배송 통로 요청서로 바로 이어지는 연결.', options: { breakLine: true } },
    { text: '미소는 1,000개 이상 서비스(2025.9), 직방은 삼성SDS 홈IoT 인수(2022)로 서비스 묶음·IoT 보유 자체는 차별점이 아님. "조율 서비스 없음"은 공개 자료 기준.', options: { color: MUTED } },
  ], { x: 0.85, y: 4.9, w: 11.6, h: 1.35, fontSize: 12.5, paraSpaceAfter: 6 });
  source(s, '각 사 앱 소개·공식 페이지, 언론 보도, 당사 조사(2026.11.26)');

  // 4. 오늘의집 vs ROOMOVE
  s = pres.addSlide({ masterName: 'CONTENT' });
  title(s, '오늘의집과 같은 목표, 반대 방향: 중립적 일정 관리가 ROOMOVE의 자리', '가장 많이 겹치는 경쟁사 비교');
  table(s, [
    ['구분', '오늘의집', 'ROOMOVE(검토안)'],
    ['고객을 처음 만나는 곳', '인테리어 콘텐츠·커머스', '이사(시공·배송 일정 충돌)'],
    ['확장 방향', '커머스 → 시공 → 이사', '이사 → 입주 직후 → 거주 중 서비스'],
    ['날짜를 직접 잡는 범위', '"원하는날 도착" 상품 지정일 배송·부분 시공\n이사·청소·인터넷은 견적 연결', '어디서 샀든, 어느 업체와 계약했든'],
    ['돈을 버는 곳', '입점 판매 수수료·직매입·자체 브랜드·광고·시공', '업체 수수료, 입주 후 반복 서비스'],
    ['떠안는 위험', '직접 시공 고정비, 직매입·자체 브랜드 재고', '시공·재고는 떠안지 않음'],
  ], { t: { x: 0.6, y: 1.6, w: 7.9, colW: [2.0, 3.2, 2.7], fontSize: 11.5, rowH: 0.62 },
    hl: (i, j) => j === 2 ? { bold: true, color: NAVY } : null });
  stat(s, 8.9, 1.6, 3.8, '3,215.7억원', '오늘의집 2025 매출, 영업손실 147.4억원\n현금성자산·단기금융상품 약 2,429억원', BLUE);
  box(s, 8.9, 3.45, 3.8, 2.0, TINT2, TINT2);
  T(s, [
    { text: '무빙패스(2026.10)', options: { bold: true, color: ORANGE, breakLine: true } },
    { text: '3만원 1회 결제, 쿠폰·적립 중심 이사 멤버십. 혜택 기간 180일의 기간제로, 이사 직후 구매가 몰리는 시기에 맞춘 구조(당사 판단).', options: { breakLine: true } },
    { text: '서비스 이용권이 아니고 일정 조율 언급 없음.', options: { color: MUTED } },
  ], { x: 9.1, y: 3.6, w: 3.4, h: 2.6, fontSize: 12, paraSpaceAfter: 6 });
  source(s, '버킷플레이스 2025 감사보고서(DART), 바이라인네트워크·벤처스퀘어(2026.10.14) 등 언론 보도, 당사 정리');

  // 5. 해외 이사 대행 사례
  s = pres.addSlide({ masterName: 'CONTENT' });
  title(s, '사람 중심 이사 대행은 인건비로 정체하거나 소프트웨어 기능으로 편입', '해외 이사 통합 신청·대행 사례');
  [['Updater', '이사 관련 신청 일괄 처리, 제휴 업체가 수수료 부담', '이사 1건당 약 1달러(2017, 당사 계산) → 업체용 소프트웨어로 전환'],
   ['Moved', '전담 인력이 이사를 대신 챙기는 대행', '시드 320만달러(2018), 2020년 이후 공개 활동 확인 못 함'],
   ['MoveEasy', '전담 인력형 이사 대행', '2024.10 AppFolio가 7,850만달러에 인수, 소프트웨어 부가 기능으로'],
   ['Hello Alfred', '전담 인력 가정 관리', '모든 세대 전담 방식의 확장 한계, 2019 건물 단위 서비스로 전환']].forEach((c, i) => {
    const x = 0.6 + i * 3.05;
    box(s, x, 1.65, 2.85, 2.9, 'F7F9FC', LINE);
    T(s, c[0], { x: x + 0.2, y: 1.8, w: 2.45, h: 0.45, fontSize: 17, bold: true, color: NAVY });
    T(s, c[1], { x: x + 0.2, y: 2.3, w: 2.45, h: 0.9, fontSize: 12 });
    T(s, c[2], { x: x + 0.2, y: 3.25, w: 2.45, h: 1.2, fontSize: 11.5, color: ORANGE });
  });
  box(s, 0.6, 4.8, 12.1, 1.6, TINT, TINT);
  T(s, '살아남은 모델의 공통점', { x: 0.85, y: 4.9, w: 5, h: 0.35, fontSize: 14, bold: true, color: BLUE });
  ['고객 무료', '계약 시 업체가 수수료', '중개사·관리회사로 대량 고객 확보', '사람 업무의 자동화'].forEach((t, i) => {
    const x = 0.85 + i * 2.95;
    s.addShape('ellipse', { x, y: 5.45, w: 0.4, h: 0.4, fill: { color: NAVY }, line: { color: NAVY } });
    T(s, String(i + 1), { x, y: 5.45, w: 0.4, h: 0.4, fontSize: 12, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    T(s, t, { x: x + 0.5, y: 5.4, w: 2.35, h: 0.5, fontSize: 12.5, bold: true, color: NAVY, valign: 'middle' });
  });
  source(s, 'AppFolio 2024 10-K(SEC), Updater 호주증권거래소 공시(2018), TechCrunch(2018.8·2019.3), 당사 정리');

  // 6. 확장 논리 연역
  s = pres.addSlide({ masterName: 'CONTENT' });
  title(s, '이사 날짜를 맡는 곳이 계약 이후 절차와 입주 후 지출을 이음', '확장 논리(연역 구조)');
  const ded = [
    ['대전제', '주거 이벤트에서 고객을 먼저 만나고 집 관련 지출을 잇는 플랫폼은 거래가 줄어도 매출을 지킴', '贝壳 2025 인테리어·임대가 중개 감소를 메워 총매출 +1.2%', '공시(해외 1곳)'],
    ['소전제 1', '이동 뒤 늘어나는 지출은 집 관련 품목과 거래 절차에 몰림', '미국 주택 구매 전후 약 3,700달러 추가, 집과 무관한 소비는 그대로(NBER)', '학술(미국)'],
    ['소전제 2', '한국은 주거 이동 사건이 큼', '매매 72.6만·전월세 279.2만건, 최근 2년 안 이사 가구 31.2%', '공식 통계'],
    ['소전제 3', '계약 뒤 절차는 잔금·입주일 전후에 몰리나 하나로 잇는 곳이 없음', '등기 잔금 뒤 60일, 전입 14일(법령), 전자계약도 외부 링크', '공식 자료'],
    ['소전제 4', '일정을 맞추려면 평면·보유 가구·입주일을 받으므로 집 단위 데이터가 쌓임', '서비스 구조에서 나오는 논리', '당사 추정'],
    ['결론', '이사 일정 조율로 들어가 계약 이후 절차와 입주 후 지출을 잇는 플랫폼으로 확장', '함께 쓰는 비율은 측정 전', '가설'],
  ];
  table(s, [['구분', '명제', '근거', '근거 강도'], ...ded], { t: { x: 0.6, y: 1.6, w: 12.1, colW: [1.3, 5.0, 4.3, 1.5], fontSize: 11.5, rowH: 0.62 },
    hl: (i, j) => i === 5 ? { bold: true, color: ORANGE, fill: { color: TINT2 } } : (j === 0 ? { bold: true, color: NAVY } : null) });
  source(s, 'KE Holdings 2025 연간실적, NBER WP 23570, 국토교통부 주택통계·2025년도 주거실태조사, 국가법령정보센터, 당사 정리');

  // 7. 국내 빈자리 (절차 흐름)
  s = pres.addSlide({ masterName: 'CONTENT' });
  title(s, '계약부터 입주 후까지 끝까지 잇는 곳은 공개 자료로 확인되지 않음', '계약 뒤 따라붙는 절차와 국내 플레이어');
  const flow = [['계약', '국토부 전자계약\n직방·호갱노노'], ['대출', '네이버페이 등\n금융 앱'], ['등기', '법무통\n(잔금 뒤 60일)'], ['전입', '정부24\n(전입 후 14일)'], ['이사·시공·배송', '업체마다 따로 예약\n(법정 기한 없음)'], ['입주 후', '아파트아이·홈닉\n정기 청소 업체']];
  flow.forEach((f, i) => {
    const x = 0.6 + i * 2.05;
    box(s, x, 1.7, 1.8, 0.65, i === 4 ? ORANGE : NAVY);
    T(s, f[0], { x, y: 1.7, w: 1.8, h: 0.65, fontSize: 14, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    T(s, f[1], { x, y: 2.45, w: 1.8, h: 0.75, fontSize: 11, align: 'center', color: INK });
    if (i < 5) arrowR(s, x + 1.82, 2.02, 0.2);
  });
  T(s, '단계마다 단일 기능 앱, 이음새는 비어 있음', { x: 0.6, y: 3.3, w: 12.1, h: 0.35, fontSize: 13, bold: true, color: ORANGE, align: 'center' });
  [['공공도 하나로 처리하지 않음', '전자계약시스템은 등기는 인터넷등기소 안내, 전입은 정부24 바로가기, 대출은 기관 메뉴로 넘김(2025년 50만7,431건, 활용률 12.04%)'],
   ['양쪽에서 다가오는 경쟁자', '홈닉은 "계약부터 거주까지 통합 플랫폼" 목표, 단 제휴 신축 단지 한정. 직방은 계약 전후 위험 관리·다음 이사까지'],
   ['ROOMOVE의 위치', '출발점은 계약도 상품도 아닌 "이사 날짜". 신축·기존 주택, 매매·전월세를 가리지 않고 잔금·입주일 전후의 실행 단계를 맡음(당사 제안)']].forEach((c, i) => {
    const x = 0.6 + i * 4.1;
    box(s, x, 3.85, 3.9, 1.75, i === 2 ? TINT : 'F7F9FC', i === 2 ? TINT : LINE);
    T(s, c[0], { x: x + 0.2, y: 4.0, w: 3.5, h: 0.4, fontSize: 14, bold: true, color: i === 2 ? BLUE : NAVY });
    T(s, c[1], { x: x + 0.2, y: 4.45, w: 3.5, h: 1.1, fontSize: 12 });
  });
  source(s, '국토부 발표(2026.1.22) 인용 보도, 직방 2025 감사보고서, 삼성물산 보도자료(2026.5·11), 국가법령정보센터, 각 앱스토어 설명(2026.11.26 조회), 당사 정리');

  // 8. 해외 선례
  s = pres.addSlide({ masterName: 'CONTENT' });
  title(s, '확장이 성립한 곳은 연 갱신 상품과 싼 고객 확보 통로를 함께 갖춤', '주택 거래·이사를 출발점으로 확장한 해외·국내 사례');
  table(s, [
    ['사례', '확장 경로', '결과'],
    ['贝壳 KE(중국)', '중개 → 인테리어·가구 → 임대관리', '성립. 두 사업 비중 약 39%'],
    ['Frontdoor(미국)', '주택 거래 → 주택 보증 연 갱신', '성립. 2025 매출 20.9억달러, 유지율 79%'],
    ['Porch(미국)', '주택 검사 소프트웨어 → 이사·보증·보험', '성립(보험). 위험 직접 질 때 큰 손실'],
    ['Zillow(미국)', '매물 광고 → 임대·모기지', '진행 중. 광고가 매출 66%'],
    ['Updater(미국)', '이사 연결 → 업체용 소프트웨어', '방향 전환. 이사 1건당 약 1달러'],
    ['Zillow Offers', '집 직접 매입', '실패. 2021 재고 감액 4.1억달러'],
    ['직방(국내)', '중개법인 직접 사업화', '구조 정비. 2025 순손실 646.2억원'],
  ], { t: { x: 0.6, y: 1.6, w: 7.6, colW: [1.9, 3.0, 2.7], fontSize: 11, rowH: 0.5 },
    hl: (i, j) => j === 2 && i < 3 ? { color: GREEN, bold: true } : (j === 2 && i >= 5 ? { color: ORANGE, bold: true } : null) });
  box(s, 8.5, 1.6, 4.2, 2.3, TINT, TINT);
  T(s, [{ text: '성립 조건', options: { bold: true, color: GREEN, breakLine: true } },
    { text: '매년 다시 돈을 내는 상품', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '업무 도구·제휴사로 싼 고객 확보', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '위험은 남에게, 수익은 수수료로', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '입주일·평면을 미리 아는 위치', options: { bullet: { indent: 14 } } }], { x: 8.7, y: 1.72, w: 3.9, h: 2.1, fontSize: 12.5, paraSpaceAfter: 3 });
  box(s, 8.5, 4.1, 4.2, 2.3, TINT2, TINT2);
  T(s, [{ text: '실패 조건', options: { bold: true, color: ORANGE, breakLine: true } },
    { text: '건당 수익이 작은 연결 모델', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '위험·고정비를 직접 떠안음', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '기존 사업자와 정면 충돌', options: { bullet: { indent: 14 } } }], { x: 8.7, y: 4.22, w: 3.9, h: 2.1, fontSize: 12.5, paraSpaceAfter: 3 });
  source(s, 'KE Holdings 연간실적, Zillow FY2021·FY2025 10-K, Frontdoor·Porch 공시(SEC), Updater 공시(2018), 직방 2025 감사보고서, 당사 정리');

  // 9. 확장 로드맵 (= 1장 논리)
  s = pres.addSlide({ masterName: 'CONTENT' });
  logicSlide(pres, s);

  // 10. 단계별 겹침
  s = pres.addSlide({ masterName: 'CONTENT' });
  title(s, '정면 경쟁은 1~3단계, 빈 곳은 업종 사이 일정 재조정과 홈닉 밖 기존 주택', '확장 단계별로 겹치는 경쟁사와 대응');
  table(s, [
    ['단계', '겹치는 곳', '겹침', '대응(당사 제안)'],
    ['1. 이사·입주 일정', '오늘의집, 홈닉, 짐싸·이사타임·헬로입주, 건설사 앱 / 계약 뒤: 직방·전자계약·금융 앱·법무통', '높음', '홈닉이 닿지 않는 기존 주택·협약 밖 단지부터. 계약 뒤 절차는 안내만'],
    ['2. 입주 직후 홈서비스', '미소(1,000개 이상 서비스), 숨고·오늘의집, 헬로입주', '높음', '서비스 수가 아니라 이사 일정 안에서 순서대로 잇는 연결로 차별화'],
    ['3. 거주 중 반복·자산 기록', '청소연구소·미소, 건설사 앱, 제조사 AS·구독, 정수기 렌탈', '높음', '경쟁보다 제휴. 어디서 샀든 전체 목록과 서비스 이력으로 차별화'],
    ['4. 주거 보장·IoT', '직방(삼성SDS 홈IoT 인수), 제조사 SmartThings·ThinQ, 하이마트 가전수리올케어', '장기', '기기는 만들지 않고 연동, 보장 위험은 보험사가 짐'],
    ['5. 다음 주거 이벤트', '직방, 홈닉 확대, 건설사·관리 앱', '제휴=경쟁', '홈닉 미도입 건설사·관리 앱·중개사무소 먼저. 도입 단지 수 분기 추적'],
  ], { t: { x: 0.6, y: 1.6, w: 12.1, colW: [2.3, 4.0, 1.2, 4.6], fontSize: 11.5, rowH: 0.75 },
    hl: (i, j) => j === 2 ? { bold: true, align: 'center', color: (i <= 2) ? ORANGE : NAVY } : (j === 0 ? { bold: true, color: NAVY } : null) });
  source(s, '2장 경쟁 분석, 각 앱스토어 설명·보도(2026.11.26 조회), SBS Biz·한국경제(하이마트, 2026.10), LG전자 가전 구독 안내, 당사 정리');

  // 11. 장기: 주거 보장·IoT
  s = pres.addSlide({ masterName: 'CONTENT' });
  title(s, '장기 방향: 쌓인 집 기록으로 주거 생애 전체를 보장하는 회사', '4단계 주거 보장·IoT(장기 검토안, 매출 추정에 넣지 않음)');
  const lad = [['① 연결', '보험대리점 제휴로 이사 시점 보증금·누수 배상·가재도구 안내, 하자담보기간·가전 보증 만료 알림'],
    ['② 기록 축적', '고장·수리·파손 기록을 가구·가전 목록에 붙여 쌓음'],
    ['③ 주택 보증', '쌓인 기록으로 보험사와 함께 설계한 보증 판매. 위험 인수는 보험사, ROOMOVE는 가격 근거 데이터와 수리 업체 연결'],
    ['④ IoT 연결', '삼성 SmartThings·LG ThinQ 등 제조사 연결 서비스, 건설사 홈 IoT와 연동해 고장 신호로 점검 시점 선제안']];
  lad.forEach((l, i) => {
    const y = 1.6 + i * 1.0;
    box(s, 0.6 + i * 0.35, y, 1.7, 0.8, i === 3 ? ORANGE : BLUE);
    T(s, l[0], { x: 0.6 + i * 0.35, y, w: 1.7, h: 0.8, fontSize: 14, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    T(s, l[1], { x: 2.5 + i * 0.35, y, w: 5.6 - i * 0.35, h: 0.8, fontSize: 11.5, valign: 'middle' });
  });
  box(s, 8.6, 1.6, 4.1, 2.3, 'F7F9FC', LINE);
  T(s, [{ text: '지금 바로는 어려운 이유', options: { bold: true, color: ORANGE, breakLine: true } },
    { text: '고객 다수가 세입자, 큰 수리는 임대인 몫(민법 623조)', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '신축은 법정 하자담보기간이 보증(마감 2년·설비 3년·방수 5년·구조 10년)', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '제3자가 파는 보증은 보험으로 볼 소지', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '대출·입주 때 보험을 요구하는 관행 없음(2013 금감원 시정)', options: { bullet: { indent: 14 } } }], { x: 8.8, y: 1.7, w: 3.8, h: 2.15, fontSize: 11, paraSpaceAfter: 2 });
  box(s, 8.6, 4.1, 4.1, 2.3, TINT, TINT);
  T(s, [{ text: '해외 선례와 한계', options: { bold: true, color: BLUE, breakLine: true } },
    { text: 'Porch: 집 정보로 보험 가격, 2025 가입자 소유 상호보험으로 전환', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: 'Frontdoor: 보증 매출 약 76%가 갱신', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '강제 없는 주택 보증은 미국 가구의 약 4%, 거래 때 받은 고객 첫해 갱신 29%', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '직방은 홈IoT 인수 뒤 경영 악화 원인으로 지목(보도)', options: { bullet: { indent: 14 } } }], { x: 8.8, y: 4.2, w: 3.8, h: 2.15, fontSize: 11, paraSpaceAfter: 2 });
  T(s, '확장 문장: "한국에는 집 보험이 팔리는 거래 순간이 없고, ROOMOVE가 그 순간을 이사 일정 안에 만든다"(당사 판단)', { x: 0.6, y: 5.75, w: 7.7, h: 0.65, fontSize: 12, bold: true, color: NAVY });
  source(s, '민법·공동주택관리법 시행령(법령), 경향신문(2013.8.13), Porch 10-K(FY2025), Frontdoor 10-K(FY2023·FY2025)·2025 연차보고서, 당사 판단');

  // 12. 데이터 해자
  s = pres.addSlide({ masterName: 'CONTENT' });
  title(s, '데이터 해자는 기본값이 아님: 조건별로 측정해 성립한 것만 주장', '해자 후보: 구매처를 묻지 않는 가구·가전 목록 + 이사 실행 기록 + 다시 쓰는 전환 비용');
  table(s, [
    ['조건(HBR 2020)', 'ROOMOVE 데이터', '현재 판단', '확인 방법'],
    ['가치가 큼', '가구·가전 치수와 새 집 평면으로 견적·시공 치수·배송 통로를 미리 맞춤', '가설', '시뮬레이션 사용·미사용 고객 비교'],
    ['가치가 오래감', '다음 이사·교체 때 다시 쓰임', '부분 성립(임차 거주 3.7년)', '입주 후 목록 갱신률'],
    ['독점적', '구매처 무관 보유 물건 전체 + 실행 기록', '같은 데이터 가진 곳 확인 못 함', '경쟁사 앱 직접 확인'],
    ['베끼기 어려움', '단지 제약·업체 지연 기록은 실제 이사를 해야 쌓임', '가설(목록 초안은 사진 도구로 쉬움)', '단지·업체 기록 누적량'],
    ['쌓일수록 빨리 좋아짐', '실행 기록이 늘수록 견적·일정 예측 정확도 상승', '가설(가치가 빨리 포화될 수 있음)', '고객 수 대비 견적 오차'],
  ], { t: { x: 0.6, y: 1.6, w: 12.1, colW: [2.0, 4.6, 2.9, 2.6], fontSize: 11.5, rowH: 0.62 },
    hl: (i, j) => j === 0 ? { bold: true, color: NAVY } : (j === 2 ? { color: ORANGE } : null) });
  T(s, '평면·3D 배치는 해자가 되기 어려움: 평면은 건설사가 먼저 갖고(홈닉), 1,400만 개 공간을 스캔한 Matterport도 CoStar에 흡수(2025.2)', { x: 0.6, y: 5.75, w: 12.1, h: 0.6, fontSize: 12, color: INK });
  source(s, 'Hagiu·Wright, HBR(2020.1), a16z(2019), CoStar 8-K(SEC), 2025년도 주거실태조사, 당사 정리');

  // 13. 생애가치 (네이티브 차트)
  s = pres.addSlide({ masterName: 'CONTENT' });
  title(s, '5년 누적 수수료는 거주 중 서비스가 붙느냐에 달림', '고객 생애가치(당사 추정, 측정 전 크기 가늠용)');
  s.addChart(pres.charts.BAR, [
    { name: '범위 하단', labels: ['거주 중 서비스 안 씀', '거주 중 서비스 씀'], values: [22, 98] },
    { name: '범위 상단', labels: ['거주 중 서비스 안 씀', '거주 중 서비스 씀'], values: [33, 113] },
  ], { x: 0.6, y: 1.55, w: 6.6, h: 4.9, barDir: 'col', barGrouping: 'clustered', chartColors: [BLUE, ORANGE],
    showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 12, dataLabelColor: INK, dataLabelFontFace: F, dataLabelFormatCode: '0"만원"',
    catAxisLabelFontFace: F, catAxisLabelFontSize: 12, catAxisLabelColor: INK, valAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
    showLegend: true, legendPos: 'b', legendFontFace: F, legendFontSize: 11, legendColor: INK,
    showTitle: true, title: '고객 1명 5년 누적 수수료 범위(만원)', titleFontFace: F, titleFontSize: 13, titleColor: NAVY });
  table(s, [
    ['항목', '안 씀', '씀'],
    ['이사 + 입주청소(1회)', '12만~21.5만', '12만~21.5만'],
    ['에어컨 이전설치(2대)', '3.6만', '3.6만'],
    ['가구·가전 연계(150만×3%)', '4.5만', '4.5만'],
    ['보관(1개월)', '2.1만~3만', '2.1만~3만'],
    ['정기 청소(4년)', '0', '69.6만'],
    ['에어컨 분해 청소(4년)', '0', '6.4만~10.5만'],
    ['합계', '약 22만~33만', '약 98만~113만'],
  ], { t: { x: 7.5, y: 1.6, w: 5.2, colW: [2.4, 1.3, 1.5], fontSize: 11, rowH: 0.42 },
    hl: (i) => i === 6 ? { bold: true, color: ORANGE } : null });
  box(s, 7.5, 5.15, 5.2, 1.25, TINT, TINT);
  T(s, [{ text: '판단 기준  ', options: { bold: true, color: BLUE } }, { text: '연계 판매·재이용은 측정 전까지 0으로 두고, 첫 이사 공헌이익만으로 고객 확보 비용의 2배 이상인지 확인' }], { x: 7.7, y: 5.25, w: 4.85, h: 1.05, fontSize: 12, valign: 'middle' });
  source(s, '숨고·미소·아정당·LG전자 공개 가격(2026.11.26 조회), 아마존 제휴 수수료표, 당사 추정. 수수료율 10%·전환율·연계 구매액은 근거 없는 가정');

  // 13-1. 1년 단위경제
  s = pres.addSlide({ masterName: 'CONTENT' });
  title(s, '중개 기준 1년 단위경제: 정기 서비스 전환과 이사 수수료 방식이 좌우', '1년 단위경제 시나리오(내부 논의 가정을 중개 기준으로 재계산, 당사 가정)');
  s.addChart(pres.charts.BAR, [{ name: '생애가치 ÷ 고객 확보 비용', labels: ['보수', '기준', '공격'], values: [0.55, 1.34, 2.53] }],
    { x: 0.6, y: 1.55, w: 5.2, h: 4.9, barDir: 'col', chartColors: [BLUE], showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 12, dataLabelColor: INK, dataLabelFontFace: F, dataLabelFormatCode: '0.00"배"',
      catAxisLabelFontFace: F, catAxisLabelFontSize: 12, catAxisLabelColor: INK, valAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { style: 'none' }, showLegend: false,
      showTitle: true, title: '1년 생애가치 ÷ 고객 확보 비용(배)', titleFontFace: F, titleFontSize: 13, titleColor: NAVY });
  table(s, [
    ['항목', '보수', '기준', '공격'],
    ['이사 공헌이익', '5만원', '8만원', '10만원'],
    ['고객 확보 비용', '10만원', '7만원', '5만원'],
    ['추가 이용 비율', '20%', '40%', '60%'],
    ['추가 이용 공헌이익(10%×70%)', '1.4만원', '2.1만원', '2.8만원'],
    ['연간 재구매율', '15%', '25%', '35%'],
    ['1년 생애가치', '약 5.5만원', '약 9.4만원', '약 12.7만원'],
  ], { t: { x: 6.1, y: 1.6, w: 6.6, colW: [2.7, 1.3, 1.3, 1.3], fontSize: 11, rowH: 0.38 }, hl: (i) => i === 5 ? { bold: true, color: ORANGE } : null });
  T(s, [
    { text: '원 시나리오(0.64·1.84·4.28배)는 추가 이용에 이익률 25%를 적용. 중개면 수수료(10%)를 넘을 수 없음', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '이사 10% 정률은 운송주선 허가가 전제. 제휴 정액(건당 3만원 예시)이면 기준 약 0.5배', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '정기 청소 1명 연 수수료 17.4만원: 이사 고객 약 38%가 입주 직후부터 쓰면 기준 시나리오가 2배 도달', options: { bullet: { indent: 14 } } },
  ], { x: 6.1, y: 4.45, w: 6.6, h: 2.0, fontSize: 12, paraSpaceAfter: 6 });
  source(s, '내부 논의 시나리오(2026.11.27), 아정당·숨고 이사 가격(업체 자료), 당사 계산. 모든 값은 근거 없는 가정이며 측정 대상');

  // 14. 리스크
  s = pres.addSlide({ masterName: 'CONTENT' });
  title(s, '주요 리스크 9가지와 대응', '확장 단계의 리스크');
  const rk = [
    ['고객 정보 활용 한계', '고객이 고른 업체에만 전달, 홍보 동의는 선택으로 분리'],
    ['입주 후 직거래 이탈', '일정 관리·지연 보상·하자 기록을 결제 건에만 제공'],
    ['정기 청소 기존 강자', '청소 자체가 아닌 일정·관리 묶음으로, 업체와 소개 수수료'],
    ['오늘의집·홈닉 기능 추가', '중립성을 빨리 검증, 건설사·관리 앱과는 제휴 검토'],
    ['이사·입주 물량 감소', '2026 아파트 입주 21.0만가구(−24.3%), 수도권부터 시작'],
    ['연계 판매의 낮은 효과', '보조 수익으로만, 사업성 계산에서 0으로 시작'],
    ['데이터 해자 약화', '목록을 견적·요청서·교체 제안에 실제로 써서 재입력 없앰'],
    ['계약 이후 절차 규제', '등록 플랫폼·자격자와 제휴, 대가 방식은 법률 자문 뒤'],
    ['함께 쓰는 비율 미검증', '단계별로 측정해 통과한 단계만 확장'],
  ];
  rk.forEach((r, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.6 + col * 4.1, y = 1.6 + row * 1.6;
    box(s, x, y, 3.9, 1.4, 'F7F9FC', LINE);
    T(s, r[0], { x: x + 0.2, y: y + 0.12, w: 3.5, h: 0.4, fontSize: 13.5, bold: true, color: NAVY });
    T(s, r[1], { x: x + 0.2, y: y + 0.55, w: 3.5, h: 0.8, fontSize: 11.5 });
  });
  source(s, '개인정보 보호법, 학술 자료(Karacaoglu 외 2022·Astashkina 외 2023), 버킷플레이스 감사보고서, 국가데이터처·부동산R114, 금융소비자보호법, Zillow FY2025 10-K, 당사 정리');

  // 15. 결론
  s = pres.addSlide({ masterName: 'CONTENT' });
  title(s, '먼저 측정할 항목으로 확장 순서·수익 구조·데이터 해자를 확정', '결론');
  table(s, [
    ['측정 항목', '통과 기준(당사 제안)'],
    ['정기 청소·가전 관리 전환율', '12개월 재이용 30% 이상'],
    ['업체가 받아들일 수수료율', '업체 이탈 없이 받을 수 있는 최고 요금'],
    ['오늘의집 외부 일정 관리 여부', '안 되면 중립성 차별점 유지'],
    ['이사매니저 처리 건수', '손익분기 월 15~40건 이상'],
    ['통로별 고객 확보 비용', '첫 이사 공헌이익 ÷ 2 이하'],
    ['시뮬레이션 고객 견적 정확도', '당일 추가 요금 뚜렷하게 낮음'],
    ['가구·가전 목록 갱신률', '첫 고객 집단 측정 후 설정'],
    ['자산 정보 기반 추천 전환율', '첫 고객 집단 측정 후 설정'],
    ['계약 이후 절차 함께 쓰는 비율', '첫 고객 집단 측정 후 설정'],
    ['입주 시 자산 정보 등록률', '첫 고객 집단 측정 후 설정'],
  ], { t: { x: 0.6, y: 1.6, w: 6.8, colW: [3.3, 3.5], fontSize: 11, rowH: 0.43 } });
  box(s, 7.7, 1.6, 5.0, 3.9, NAVY);
  T(s, [{ text: '성립에 필요한 네 가지 조건', options: { bold: true, color: 'AFC6E9', breakLine: true } },
    { text: '정보는 판매가 아닌 동의 범위 안의 제안 능력으로', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '누적 매출은 연계 판매가 아닌 반복 서비스로', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '대출·중개·등기는 등록 플랫폼·자격자와 제휴로', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '다음 이사는 계산에서 뺌', options: { bullet: { indent: 14 }, breakLine: true } },
    { text: '장기 목표: 이사를 입구로 주거 생애주기를 잇는 플랫폼', options: { bullet: { indent: 14 }, bold: true, color: 'AFC6E9' } }], { x: 7.95, y: 1.85, w: 4.5, h: 3.5, fontSize: 14, color: WHITE, paraSpaceAfter: 10 });
  source(s, '리서치/확장(2026.11.25), 기존 보고서(v1) 5장, 당사 정리');
}

(async () => {
  let p = new pptxgen(); setup(p);
  let s = p.addSlide({ masterName: 'CONTENT' }); logicSlide(p, s);
  await p.writeFile({ fileName: path.join(OUT, 'logic.pptx') });
  p = new pptxgen(); setup(p); fullDeck(p);
  await p.writeFile({ fileName: path.join(OUT, 'full.pptx') });
  console.log('done');
})();
