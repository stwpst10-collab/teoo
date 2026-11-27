// ROOMOVE 확장·경쟁 분석 발표: 팀 장표(Figma ROOMOVE_Drafts) 톤 판
const pptxgen = require('pptxgenjs');
const OUT = process.argv[2];
const F = 'Malgun Gothic';
const INK = '111111', SUB = '444444', MUTE = '8A8A8A', CARD = 'F2F2F2', BLACK = '161616', WHITE = 'FFFFFF', LINE = 'DDDDDD';
const V = ['D9D7FF', 'B9B6FF', '8E8AFF', '5F5BF7', '3B37E8'];
const ACC = '3B37E8', RED = 'D4380D';

const p = new pptxgen();
p.layout = 'LAYOUT_WIDE';
p.theme = { headFontFace: F, bodyFontFace: F };
let s;
const T = (text, o) => s.addText(text, Object.assign({ isTextBox: true, fontFace: F, color: INK, margin: 0, valign: 'top' }, o));
const rect = (x, y, w, h, c) => s.addShape('rect', { x, y, w, h, fill: { color: c }, line: { color: c } });
const pill = (x, y, w, h, c) => s.addShape('roundRect', { x, y, w, h, fill: { color: c }, line: { color: c }, rectRadius: h / 2 });
function page(label, title, sub) {
  s = p.addSlide(); s.background = { color: WHITE };
  T(label, { x: 0.6, y: 0.45, w: 8, h: 0.3, fontSize: 13, bold: true });
  T(title, { x: 0.6, y: 0.8, w: 12.1, h: 0.6, fontSize: 26, bold: true });
  if (sub) T(sub, { x: 0.6, y: 1.42, w: 12.1, h: 0.35, fontSize: 14, bold: true, color: '333333' });
}
function divider(text) { s = p.addSlide(); s.background = { color: WHITE }; T(text, { x: 0.8, y: 3.2, w: 11, h: 1, fontSize: 40, bold: true }); }
function foot(text) { T(text, { x: 0.6, y: 6.85, w: 12.13, h: 0.3, fontSize: 9, color: MUTE, align: 'center' }); }
function tbl(rows, opt) {
  const head = rows[0].map(t => ({ text: t, options: { bold: true, color: WHITE, fill: { color: BLACK } } }));
  const body = rows.slice(1).map((r, i) => r.map((t, j) => ({ text: t, options: Object.assign({ fill: { color: i % 2 ? WHITE : 'F6F6F6' }, bold: j === 0 }, (opt.hl && opt.hl(i, j)) || {}) })));
  s.addTable([head, ...body], Object.assign({ fontFace: F, fontSize: 12, color: INK, border: { type: 'solid', color: WHITE, pt: 2 }, valign: 'middle', margin: [4, 10, 4, 10] }, opt.t));
}

// 1. 표지
s = p.addSlide(); s.background = { color: WHITE };
T('사내벤처 · 2026.11', { x: 0.8, y: 2.3, w: 8, h: 0.35, fontSize: 14, bold: true, color: MUTE });
T('ROOMOVE 확장·경쟁 분석', { x: 0.8, y: 2.75, w: 11.5, h: 0.9, fontSize: 40, bold: true });
T('이사에서 시작해, 주거 라이프사이클 전체로', { x: 0.8, y: 3.75, w: 11.5, h: 0.5, fontSize: 20, bold: true, color: ACC });

// 2. 요약
page('Overview', '이사는 지출과 절차가 한 날짜로 몰리는 순간', '그 날짜를 맡는 곳이 다음 지출을 잇는다.');
[['351.8만건', '2025 주택 매매 72.6만 +\n전월세 279.2만건'], ['약 39%', '중국 贝壳 인테리어·임대\n매출 비중(2025)'], ['약 2배', '미국 주택 구매 이사 가구의\n첫해 집 관련 지출'], ['98만~113만원', '고객 1명 5년 누적 수수료\n(거주 중 서비스 이용 시, 당사 추정)']].forEach((r, i) => {
  const x = 0.6 + i * 3.08;
  rect(x, 2.2, 2.9, 2.3, CARD);
  T(r[0], { x: x + 0.25, y: 2.45, w: 2.5, h: 0.8, fontSize: i === 3 ? 25 : 30, bold: true, color: i === 3 ? RED : INK });
  T(r[1], { x: x + 0.25, y: 3.4, w: 2.5, h: 0.9, fontSize: 12, color: SUB });
});
rect(0.6, 4.8, 12.13, 1.5, BLACK);
T([{ text: '이사 회사가 아니라, 주거 생애주기 플랫폼', options: { bold: true, fontSize: 17, breakLine: true } },
   { text: '이사로 고객과 집 정보를 함께 얻고, 서비스가 생길 때마다 그 집의 이력을 쌓는다. 요청에 답하는 대신 필요한 서비스를 먼저 제안하는 것이 목표다(가설).', options: { fontSize: 12.5, color: 'CFCFCF' } }],
  { x: 0.9, y: 5.05, w: 11.5, h: 1.1, color: WHITE, paraSpaceAfter: 6 });
foot('자료: 국토교통부 주택통계(2026.1), KE Holdings 2025 연간실적, NAHB(2026.8), 당사 추정');

// 3. 경쟁
divider('1. 경쟁 구도');
page('Competition', '견적 연결은 이미 흔하고, 업종 사이 조율은 비어 있다', '여러 업체 일정을 순서대로 맞추고, 밀리면 다시 맞춰 주는 곳은 공개 자료로 확인되지 않는다.');
const fn = [['여러 업종 견적·신청', '짐싸 · 이사타임 · 오늘의집 · 헬로입주 · 홈닉 · 미소', '이미 흔함'], ['도면 위 가구 배치', '오늘의집 3D · 아키스케치 · 한샘·홈닉', '도구 존재'], ['한 회사 안 공정 관리', '삼성 인테리어핏 · IKEA 주방 · 오늘의집 직접 시공', '자사 상품 한정'], ['업종 사이 일정 재조정', '확인된 곳 없음', 'ROOMOVE 자리']];
fn.forEach((r, i) => {
  const y = 2.25 + i * 0.95, last = i === 3;
  T(r[0], { x: 0.6, y, w: 3.2, h: 0.6, fontSize: 14, bold: true, valign: 'middle' });
  pill(3.9, y + 0.06, 6.3, 0.5, last ? BLACK : CARD);
  T(r[1], { x: 3.9, y: y + 0.06, w: 6.3, h: 0.5, fontSize: 12, bold: true, align: 'center', valign: 'middle', color: last ? WHITE : SUB });
  T(r[2], { x: 10.4, y, w: 2.3, h: 0.6, fontSize: 13, bold: true, valign: 'middle', color: last ? ACC : MUTE });
});
foot('미소 1,000개 이상 서비스(2025.9), 직방 삼성SDS 홈IoT 인수(2022): 서비스 수·IoT 보유 자체는 차별점이 아니다. 자료: 각 사 앱 소개·보도(2026.11)');

page('Competition', '오늘의집과 같은 목표, 반대 방향', '오늘의집은 커머스 고객을 시공·이사로, ROOMOVE는 이사 고객을 주거 생활로 넓힌다.');
const cmp = [['처음 만나는 곳', '인테리어 콘텐츠·커머스', '이사(시공·배송 일정 충돌)'], ['날짜를 직접 잡는 범위', '자사 지정일 배송·부분 시공', '어디서 샀든, 어느 업체든'], ['돈을 버는 곳', '입점 수수료·직매입·광고·시공', '업체 수수료, 거주 중 반복 서비스'], ['떠안는 위험', '시공 고정비, 재고', '시공·재고는 떠안지 않음']];
rect(3.5, 2.2, 4.4, 4.3, CARD); rect(8.2, 2.2, 4.5, 4.3, BLACK);
T('오늘의집', { x: 3.75, y: 2.4, w: 4, h: 0.4, fontSize: 16, bold: true });
T('ROOMOVE', { x: 8.45, y: 2.4, w: 4, h: 0.4, fontSize: 16, bold: true, color: WHITE });
cmp.forEach((r, i) => {
  const y = 3.05 + i * 0.85;
  T(r[0], { x: 0.6, y, w: 2.8, h: 0.6, fontSize: 13, bold: true, color: SUB, valign: 'middle' });
  T(r[1], { x: 3.75, y, w: 4, h: 0.6, fontSize: 13, valign: 'middle' });
  T(r[2], { x: 8.45, y, w: 4.1, h: 0.6, fontSize: 13, bold: true, valign: 'middle', color: WHITE });
});
foot('오늘의집 2025 매출 3,215.7억원·영업손실 147.4억원. 무빙패스(2026.10)는 180일 쿠폰·적립 멤버십으로 일정 조율 언급 없음. 자료: 감사보고서(DART), 언론 보도');

page('Competition', '사람이 대신 챙기는 이사 대행은 커지지 못했다', '살아남은 곳은 고객 무료·업체 수수료·제휴로 고객 확보·업무 자동화를 갖췄다.');
[['Updater', '이사 1건당 약 1달러 →\n업체용 소프트웨어로 전환'], ['Moved', '전담 인력 대행,\n2020년 이후 활동 확인 못 함'], ['MoveEasy', '2024 AppFolio가 인수,\n소프트웨어 부가 기능으로'], ['Hello Alfred', '세대별 전담 인력의 한계,\n건물 단위 서비스로 전환']].forEach((c, i) => {
  const x = 0.6 + i * 3.08;
  rect(x, 2.25, 2.9, 2.5, i === 0 ? BLACK : CARD);
  T(c[0], { x: x + 0.25, y: 2.5, w: 2.5, h: 0.45, fontSize: 18, bold: true, color: i === 0 ? WHITE : INK });
  T(c[1], { x: x + 0.25, y: 3.15, w: 2.5, h: 1.3, fontSize: 12.5, color: i === 0 ? 'CFCFCF' : SUB });
});
T('ROOMOVE도 사람 손은 지연 대응으로 좁히고, 중개사무소·관리사무소·입주박람회와 제휴해 고객을 얻는다.', { x: 0.6, y: 5.2, w: 12.1, h: 0.5, fontSize: 14, bold: true });
foot('자료: AppFolio 2024 10-K, Updater 호주증권거래소 공시(2018), TechCrunch(2018·2019)');

// 4. 확장
divider('2. 확장 전략');
page('Expansion', '그 날짜를 맡는 곳이 다음 지출을 잇는다', '해외 선례와 국내 빈자리에서 나온 확장 논리.');
const ded = [['대전제', '주거 이벤트에서 고객을 먼저 만나는 플랫폼은 거래가 줄어도 매출을 지킨다', '贝壳 2025 인테리어·임대가 중개 감소를 메움'], ['소전제 1', '이사 뒤 늘어나는 지출은 집 관련 품목과 절차에 몰린다', '미국 학술(NBER)'], ['소전제 2', '한국은 주거 이동이 크다', '매매·전월세 351.8만건, 최근 2년 이사 31.2%'], ['소전제 3', '계약 뒤 절차는 잔금·입주일에 몰리지만 잇는 곳이 없다', '등기 60일·전입 14일(법령)']];
ded.forEach((r, i) => {
  const y = 2.15 + i * 0.85;
  rect(0.6, y, 12.13, 0.72, CARD);
  T(r[0], { x: 0.85, y, w: 1.4, h: 0.72, fontSize: 13, bold: true, color: ACC, valign: 'middle' });
  T(r[1], { x: 2.3, y, w: 6.6, h: 0.72, fontSize: 13, bold: true, valign: 'middle' });
  T(r[2], { x: 9.0, y, w: 3.6, h: 0.72, fontSize: 11, color: MUTE, valign: 'middle' });
});
rect(0.6, 5.65, 12.13, 0.85, BLACK);
T([{ text: '결론  ', options: { color: 'BDBDBD' } }, { text: '이사 일정 조율로 들어가 입주 뒤 집 관련 지출까지 잇는다. 함께 쓰는 비율은 측정 전 가설이다.', options: { color: WHITE } }], { x: 0.85, y: 5.65, w: 11.7, h: 0.85, fontSize: 14, bold: true, valign: 'middle' });

page('Expansion', '계약부터 입주 후까지, 끝까지 잇는 곳은 없다', '공공도 민간도 단계마다 한 가지 기능에 머문다.');
const fl = [['계약', '전자계약·직방'], ['대출', '금융 앱'], ['등기', '법무통'], ['전입', '정부24'], ['이사·시공·배송', '업체마다 따로'], ['입주 후', '관리 앱·청소 업체']];
fl.forEach((f, i) => {
  const x = 0.6 + i * 2.05;
  pill(x, 2.35, 1.85, 0.6, i === 4 ? ACC : BLACK);
  T(f[0], { x, y: 2.35, w: 1.85, h: 0.6, fontSize: 14, bold: true, color: WHITE, align: 'center', valign: 'middle' });
  T(f[1], { x, y: 3.1, w: 1.85, h: 0.35, fontSize: 11.5, color: SUB, align: 'center' });
  if (i < 5) { s.addShape('ellipse', { x: x + 1.92, y: 3.6, w: 0.5, h: 0.5, fill: { color: V[3] }, line: { color: V[3] } }); T('?', { x: x + 1.92, y: 3.6, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: WHITE, align: 'center', valign: 'middle' }); }
});
T('단계 사이 이음새는 고객이 직접 메운다.', { x: 0.6, y: 4.5, w: 12.1, h: 0.45, fontSize: 18, bold: true, align: 'center' });
T('홈닉은 "계약부터 거주까지" 목표지만 제휴 신축 단지 한정, 직방은 계약 전후에 머문다. ROOMOVE의 출발점은 계약도 상품도 아닌 "이사 날짜"다.', { x: 1.6, y: 5.15, w: 10.1, h: 0.8, fontSize: 13, color: SUB, align: 'center' });
foot('자료: 국토부 전자계약시스템, 삼성물산 보도자료(2026.5·11), 직방 감사보고서, 국가법령정보센터');

page('Expansion', '확장이 성립한 곳은 매년 다시 돈을 내는 상품을 가졌다', '거래 시점에 고객을 만나고, 위험은 남에게 넘기고, 반복 상품으로 번다.');
rect(0.6, 2.2, 5.9, 4.3, CARD); rect(6.83, 2.2, 5.9, 4.3, BLACK);
T('성립', { x: 0.85, y: 2.4, w: 5, h: 0.4, fontSize: 16, bold: true, color: ACC });
T('실패', { x: 7.08, y: 2.4, w: 5, h: 0.4, fontSize: 16, bold: true, color: 'FF8A65' });
[['贝壳', '중개 → 인테리어·임대, 매출 비중 약 39%'], ['Frontdoor', '거래 시점 → 매년 갱신 보증, 유지율 79%'], ['Porch', '검사 소프트웨어 → 보험(위험 직접 질 땐 손실)']].forEach((r, i) => {
  T(r[0], { x: 0.85, y: 3.05 + i * 1.0, w: 1.8, h: 0.4, fontSize: 14, bold: true });
  T(r[1], { x: 2.6, y: 3.05 + i * 1.0, w: 3.7, h: 0.8, fontSize: 12.5, color: SUB });
});
[['Updater', '건당 수익이 작은 연결 모델'], ['Zillow Offers', '집값 위험을 직접 떠안음, 재고 감액 4.1억달러'], ['직방', '중개법인 직접 사업화, 2025 순손실 646.2억원']].forEach((r, i) => {
  T(r[0], { x: 7.08, y: 3.05 + i * 1.0, w: 2.0, h: 0.4, fontSize: 14, bold: true, color: WHITE });
  T(r[1], { x: 9.1, y: 3.05 + i * 1.0, w: 3.5, h: 0.8, fontSize: 12.5, color: 'CFCFCF' });
});
foot('자료: KE Holdings·Zillow·Frontdoor·Porch 공시, Updater 공시(2018), 직방 감사보고서');

// 로드맵(팀 1장과 같은 구성)
page('Expansion', '이사에서 시작해, 주거 라이프사이클 전체로', '이사 때 쌓인 집 정보가 다음 서비스를 먼저 제안하게 만든다.');
const st = [['이사·입주', '업종 사이 일정 조율', '업체 수수료'], ['입주 직후', '청소·설치·인터넷·수리 연결', '성사 수수료'], ['거주 중 관리', '정기 관리·소모품 제안', '정기 이용료·제휴 수수료'], ['스마트홈·보장', '집 데이터 해상도↑', '관리 매출 → 보증 수수료'], ['다음 이사·거래', '재이사·임대·매매 지원', '제휴사 이용료']];
{ const x0 = 0.6, W = 12.13, n = st.length, w = W / n, y = 2.15, h = 0.62;
  [...st.keys()].reverse().forEach((i) => { const x = x0 + i * w; s.addShape(i === 0 ? 'homePlate' : 'chevron', { x: i === 0 ? x : x - 0.18, y, w: (i === 0 || i === n - 1) ? w + 0.18 : w + 0.36, h, fill: { color: V[i] }, line: { color: V[i] } }); });
  s.addShape('line', { x: x0, y: y + 0.97, w: W, h: 0, line: { color: 'D0D0D0', width: 1, dashType: 'dash' } });
  st.forEach((r, i) => { const x = x0 + i * w;
    T(r[0], { x: x + 0.1, y, w: w - 0.2, h, fontSize: 15, bold: true, align: 'center', valign: 'middle', color: i >= 3 ? WHITE : INK });
    if (i === 0 || i === 3) T(i === 0 ? '현재' : '장기', { x: x + 0.1, y: y - 0.32, w: w - 0.2, h: 0.28, fontSize: 10, bold: true, color: ACC, align: 'center' });
    s.addShape('ellipse', { x: x + w / 2 - 0.07, y: y + 0.9, w: 0.14, h: 0.14, fill: { color: ACC }, line: { color: ACC } });
    T(r[1], { x: x + 0.1, y: y + 1.12, w: w - 0.2, h: 0.32, fontSize: 12, bold: true, align: 'center' });
    T(r[2], { x: x + 0.1, y: y + 1.45, w: w - 0.2, h: 0.3, fontSize: 10.5, color: MUTE, align: 'center' }); }); }
rect(0.6, 4.25, 5.9, 2.25, CARD);
T('이사 때 쌓이는 집 정보', { x: 0.85, y: 4.5, w: 5.4, h: 0.4, fontSize: 16, bold: true });
T('어디서 샀든, 어느 업체든 우리 집 데이터로 쌓인다', { x: 0.85, y: 4.93, w: 5.4, h: 0.3, fontSize: 11.5, color: SUB });
['새집 평면·치수', '가구·가전 목록과 구매 연도', '업체별 일정 이행 기록', '고장·수리·교체 이력'].forEach((t, i) => T(t, { x: 0.85 + (i % 2) * 2.7, y: 5.4 + Math.floor(i / 2) * 0.42, w: 2.6, h: 0.35, fontSize: 12.5, bold: true }));
rect(6.83, 4.25, 5.9, 2.25, BLACK);
T('그래서', { x: 7.08, y: 4.5, w: 5.4, h: 0.3, fontSize: 11.5, bold: true, color: 'BDBDBD' });
T('고객이 늘수록 기록이 쌓이고,\n기록이 쌓일수록 제안이 정확해진다.', { x: 7.08, y: 4.83, w: 5.4, h: 0.8, fontSize: 17, bold: true, color: WHITE });
T([{ text: '먼저 확인할 것  ', options: { bold: true, color: 'BDBDBD' } }, { text: '정기 서비스 전환율 · 고객 확보 비용 · 집 정보 등록률', options: { color: WHITE } }], { x: 7.08, y: 5.85, w: 5.4, h: 0.45, fontSize: 11.5 });

page('Expansion', '정면 경쟁은 1~3단계, 빈 곳은 업종 사이 일정과 홈닉 밖 기존 주택', '단계마다 겹치는 곳과 대응.');
tbl([['단계', '겹치는 곳', '대응'],
  ['이사·입주', '오늘의집 · 홈닉 · 짐싸 · 헬로입주', '홈닉이 닿지 않는 기존 주택·협약 밖 단지부터'],
  ['입주 직후', '미소(1,000개 이상 서비스) · 숨고 · 오늘의집', '서비스 수가 아니라 이사 일정 안 순서대로 잇기'],
  ['거주 중 관리', '청소연구소 · 미소 · 제조사 AS·구독 · 정수기 렌탈', '경쟁보다 제휴, 어디서 샀든 전체 목록으로'],
  ['스마트홈·보장', '직방(삼성SDS 홈IoT) · 제조사 SmartThings·ThinQ · 하이마트', '기기는 만들지 않고 연동, 위험은 보험사가'],
  ['다음 이사·거래', '직방 · 홈닉 확대 · 건설사·관리 앱', '홈닉 미도입 건설사·관리 앱·중개사무소 먼저']],
  { t: { x: 0.6, y: 2.15, w: 12.13, colW: [2.0, 5.0, 5.13], rowH: 0.68 } });
foot('자료: 각 사 앱스토어 설명·보도(2026.11), 당사 정리');

page('Expansion', '장기: 쌓인 집 기록으로 주거 생애 전체를 보장한다', '지금은 보장을 직접 팔지 않고, 연결과 기록부터 쌓는다.');
[['연결', '보험대리점 제휴로 이사 때 보증금·누수·가재도구 안내, 하자 보수 기한 알림'], ['기록 축적', '고장·수리·파손 기록을 가구·가전 목록에 붙인다'], ['주택 보증', '쌓인 기록으로 보험사와 함께 보증 상품. 위험은 보험사가 진다'], ['IoT 연결', '제조사·건설사 IoT와 연동해 고장 신호로 점검 시점을 먼저 제안']].forEach((r, i) => {
  const y = 2.15 + i * 1.0;
  pill(0.6, y + 0.12, 1.9, 0.56, V[i + 1]);
  T(r[0], { x: 0.6, y: y + 0.12, w: 1.9, h: 0.56, fontSize: 14, bold: true, align: 'center', valign: 'middle', color: i >= 2 ? WHITE : INK });
  T(r[1], { x: 2.75, y: y + 0.12, w: 4.6, h: 0.75, fontSize: 12.5, valign: 'middle' });
});
rect(7.7, 2.15, 5.03, 4.35, BLACK);
T('지금 바로 어려운 이유', { x: 7.95, y: 2.4, w: 4.6, h: 0.4, fontSize: 15, bold: true, color: WHITE });
T([{ text: '세입자의 큰 수리는 집주인 몫(민법 623조)', options: { bullet: { indent: 14 }, breakLine: true } },
   { text: '신축은 법정 하자 보수 기간(마감 2년·설비 3년·구조 10년)이 보증', options: { bullet: { indent: 14 }, breakLine: true } },
   { text: '제3자가 파는 보증은 보험으로 볼 소지', options: { bullet: { indent: 14 }, breakLine: true } },
   { text: '미국도 강제 없는 주택 보증은 가구의 약 4%', options: { bullet: { indent: 14 }, breakLine: true } },
   { text: '직방은 홈IoT 인수 뒤 경영 악화 원인으로 지목(보도)', options: { bullet: { indent: 14 } } }],
  { x: 7.95, y: 2.95, w: 4.6, h: 3.4, fontSize: 12.5, color: 'E0E0E0', paraSpaceAfter: 8 });
foot('자료: 민법·공동주택관리법 시행령, Frontdoor·Porch 공시, 언론 보도. 매출 추정에 넣지 않는 장기 검토안');

page('Expansion', '데이터 해자는 기본값이 아니다', '해자 후보는 어디서 샀든 집 안 가구·가전 목록과 업체별 이행 기록. 조건마다 측정해 성립한 것만 주장한다.');
tbl([['조건', '현재 판단', '확인 방법'],
  ['가치가 큰가', '가설: 이사 당일 추가 요금·일정 변경이 줄어야 성립', '집 정보 사용·미사용 고객 비교'],
  ['오래가는가', '부분 성립: 임차 평균 거주 3.7년이라 재사용 주기가 김', '입주 후 목록 갱신률'],
  ['독점적인가', '같은 데이터를 가진 곳 확인 못 함', '경쟁사 앱 직접 확인'],
  ['베끼기 어려운가', '가설: 목록 초안은 사진 도구로 쉽게 만듦', '단지·업체 기록 누적량'],
  ['쌓일수록 좋아지나', '가설: 추가 데이터 가치가 빨리 포화될 수 있음', '고객 수 대비 견적 오차']],
  { t: { x: 0.6, y: 2.2, w: 12.13, colW: [2.4, 5.6, 4.13], rowH: 0.68 }, hl: (i, j) => j === 1 ? { color: SUB } : null });
foot('평면·3D 배치는 해자가 아니다: 평면은 건설사가 먼저 갖고, 1,400만 공간을 스캔한 Matterport도 CoStar에 흡수. 자료: HBR(2020), CoStar 공시');

// 5. 사업성
divider('3. 사업성');
page('Unit Economics', '5년 누적 수수료는 거주 중 서비스가 갈랐다', '정기 청소·에어컨 분해 청소가 붙으면 고객 1명의 수수료가 약 3.5~4.4배가 된다.');
rect(0.6, 2.2, 5.9, 3.0, CARD); rect(6.83, 2.2, 5.9, 3.0, BLACK);
T('거주 중 서비스 안 씀', { x: 0.9, y: 2.45, w: 5, h: 0.35, fontSize: 14, bold: true, color: SUB });
T('22만~33만원', { x: 0.9, y: 3.0, w: 5.4, h: 1.0, fontSize: 44, bold: true });
T('이사·입주청소·에어컨 이전·연계 구매·보관', { x: 0.9, y: 4.25, w: 5.4, h: 0.4, fontSize: 12, color: SUB });
T('거주 중 서비스 씀', { x: 7.13, y: 2.45, w: 5, h: 0.35, fontSize: 14, bold: true, color: 'BDBDBD' });
T('98만~113만원', { x: 7.13, y: 3.0, w: 5.4, h: 1.0, fontSize: 44, bold: true, color: 'FF8A65' });
T('+ 입주 1년 뒤부터 정기 청소 4년, 에어컨 분해 청소', { x: 7.13, y: 4.25, w: 5.4, h: 0.4, fontSize: 12, color: 'CFCFCF' });
T('판단 기준: 연계 판매·재이용은 측정 전까지 0으로 두고, 첫 이사 이익만으로 고객 확보 비용의 2배 이상인지 본다.', { x: 0.6, y: 5.55, w: 12.1, h: 0.5, fontSize: 14, bold: true });
foot('자료: 숨고·미소·아정당·LG전자 공개 가격(2026.11), 당사 추정. 수수료율 10%·전환율은 근거 없는 가정');

page('Unit Economics', '1년 기준으로는 정기 서비스 전환이 관건이다', '중개 구조에서는 추가 이용 이익이 수수료(10%)를 넘을 수 없다.');
[['보수', '0.55배', '확보 비용 10만원'], ['기준', '1.34배', '확보 비용 7만원'], ['공격', '2.53배', '확보 비용 5만원']].forEach((r, i) => {
  const x = 0.6 + i * 4.1;
  rect(x, 2.2, 3.9, 2.4, i === 1 ? BLACK : CARD);
  T(r[0], { x: x + 0.3, y: 2.4, w: 3.3, h: 0.35, fontSize: 14, bold: true, color: i === 1 ? 'BDBDBD' : SUB });
  T(r[1], { x: x + 0.3, y: 2.85, w: 3.3, h: 1.0, fontSize: 46, bold: true, color: i === 2 ? RED : (i === 1 ? WHITE : INK) });
  T('1년 생애가치 ÷ ' + r[2], { x: x + 0.3, y: 3.95, w: 3.3, h: 0.4, fontSize: 12, color: i === 1 ? 'CFCFCF' : SUB });
});
[['이사 정률 10%는 운송주선 허가가 있어야 한다', '허가 없이 건당 정액(3만원 예시)이면 기준이 약 0.5배'], ['정기 청소 1명이 연 17.4만원 수수료를 만든다', '이사 고객 약 38%가 입주 직후부터 쓰면 기준이 2배에 닿는다']].forEach((r, i) => {
  const y = 4.95 + i * 0.75;
  T(r[0], { x: 0.6, y, w: 6.2, h: 0.6, fontSize: 14, bold: true, valign: 'middle' });
  T(r[1], { x: 6.9, y, w: 5.8, h: 0.6, fontSize: 13, color: SUB, valign: 'middle' });
});
foot('내부 논의 시나리오(2026.11.27) 가정값을 중개 기준으로 재계산. 원 시나리오(0.64·1.84·4.28배)는 추가 이용 이익률 25% 적용. 모두 측정 전 가정');

// 6. 리스크·검증
divider('4. 리스크와 검증');
page('Risk', '확장 단계의 리스크 9가지', '대부분은 직접 떠안지 않고, 측정해서 통과한 단계만 넓히는 방식으로 대응한다.');
[['고객 정보 활용', '고객이 고른 업체에만 전달'], ['입주 후 직거래 이탈', '일정·지연 보상·하자 기록은 결제 건만'], ['정기 청소 기존 강자', '청소 말고 일정·관리 묶음으로'], ['경쟁사 기능 추가', '중립성을 빨리 검증, 건설사와 제휴'], ['이사·입주 물량 감소', '2026 입주 −24.3%, 수도권부터'], ['연계 판매 효과 낮음', '보조 수익, 계산에선 0으로'], ['데이터 해자 약화', '목록을 견적·요청서에 실제로 씀'], ['계약 뒤 절차 규제', '등록 플랫폼·자격자와 제휴'], ['함께 쓰는 비율 미검증', '단계별 측정 후 확장']].forEach((r, i) => {
  const c = i % 3, rr = Math.floor(i / 3), x = 0.6 + c * 4.1, y = 2.15 + rr * 1.45;
  rect(x, y, 3.9, 1.28, CARD);
  T(r[0], { x: x + 0.25, y: y + 0.2, w: 3.4, h: 0.4, fontSize: 14, bold: true });
  T(r[1], { x: x + 0.25, y: y + 0.65, w: 3.4, h: 0.5, fontSize: 12, color: SUB });
});

page('Next', '먼저 측정할 13가지', '확장 순서, 수익 구조, 데이터 해자를 이 숫자로 정한다.');
const ms = ['정기 청소·가전 관리 전환율', '업체가 받아들일 수수료율', '오늘의집 외부 일정 관리 여부', '이사매니저 처리 건수', '통로별 고객 확보 비용', '집 정보 사용 고객의 견적 정확도', '가구·가전 목록 갱신률', '집 정보 기반 추천 전환율', '계약 뒤 절차 함께 쓰는 비율', '입주 시 집 정보 등록률', '시공 기록 완성률', '수리 1회 해결률', '업체별 공종 하자 재발률'];
ms.forEach((m, i) => {
  const c = i % 2, rr = Math.floor(i / 2), x = 0.6 + c * 3.55, y = 2.15 + rr * 0.62;
  T(String(i + 1).padStart(2, '0'), { x, y, w: 0.6, h: 0.6, fontSize: 16, bold: true, color: ACC, valign: 'middle' });
  T(m, { x: x + 0.6, y, w: 2.85, h: 0.6, fontSize: 13, bold: true, valign: 'middle' });
});
rect(7.9, 2.15, 4.83, 4.35, BLACK);
T('성립에 필요한 조건', { x: 8.15, y: 2.4, w: 4.4, h: 0.4, fontSize: 15, bold: true, color: 'BDBDBD' });
T([{ text: '정보는 팔지 않고, 동의 안에서 더 정확히 제안한다', options: { bullet: { indent: 14 }, breakLine: true } },
   { text: '누적 매출은 연계 판매가 아니라 반복 서비스에서 낸다', options: { bullet: { indent: 14 }, breakLine: true } },
   { text: '대출·중개·등기는 자격자와 제휴한다', options: { bullet: { indent: 14 }, breakLine: true } },
   { text: '다음 이사는 계산에서 뺀다', options: { bullet: { indent: 14 } } }],
  { x: 8.15, y: 2.95, w: 4.4, h: 3.4, fontSize: 14, bold: true, color: WHITE, paraSpaceAfter: 12 });

p.writeFile({ fileName: OUT }).then(() => console.log('done'));
