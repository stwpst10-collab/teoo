// 4단계 장기: 주거 보장(4-A)과 스마트홈(4-B)을 같은 라인의 두 줄로 보여 주는 장표 1장
// 사용: node build_slide_two_track.js <출력.pptx> team|mine
const pptxgen = require('pptxgenjs');
const [OUT, STYLE] = [process.argv[2], process.argv[3] || 'team'];
const F = 'Malgun Gothic';
const team = STYLE === 'team';
const INK = team ? '111111' : '1F2937', SUB = '444444', MUTE = team ? '8A8A8A' : '6B7280';
const CARD = team ? 'F2F2F2' : 'F7F9FC', DARK = team ? '161616' : '15233B', WHITE = 'FFFFFF';
const ACC = team ? '3B37E8' : '2A78D6', ORG = team ? 'D4380D' : 'EB6834';
const p = new pptxgen(); p.layout = 'LAYOUT_WIDE'; p.theme = { headFontFace: F, bodyFontFace: F };
const s = p.addSlide(); s.background = { color: WHITE };
const T = (t, o) => s.addText(t, Object.assign({ isTextBox: true, fontFace: F, color: INK, margin: 0, valign: 'top' }, o));
const rect = (x, y, w, h, c) => s.addShape(team ? 'rect' : 'roundRect', Object.assign({ x, y, w, h, fill: { color: c }, line: { color: c } }, team ? {} : { rectRadius: 0.08 }));

if (team) {
  T('Expansion · 장기', { x: 0.6, y: 0.45, w: 8, h: 0.3, fontSize: 13, bold: true });
  T('보장과 스마트홈, 같은 데이터에서 두 줄로', { x: 0.6, y: 0.8, w: 12.1, h: 0.6, fontSize: 26, bold: true });
  T('둘 다 우리 집 데이터에서 출발하지만, 따로 굴러간다. 함께 돌면 더 강해진다.', { x: 0.6, y: 1.42, w: 12.1, h: 0.35, fontSize: 14, bold: true, color: '333333' });
} else {
  T('장기 방향: 보장과 스마트홈, 같은 우리 집 데이터에서 두 줄로', { x: 0.6, y: 0.35, w: 12.1, h: 0.75, fontSize: 22, bold: true, color: '15233B' });
  T('4단계(장기 검토안, 매출 추정에 넣지 않음): 둘은 따로 굴러가고, 함께 돌면 서로를 강화', { x: 0.6, y: 1.05, w: 12.1, h: 0.4, fontSize: 14, color: MUTE });
}
// 공통 출발점
rect(0.6, 1.95, 12.13, 0.62, team ? 'E9E8FF' : 'EEF4FC');
T([{ text: '공통 출발점  ', options: { bold: true, color: ACC } }, { text: '우리 집 데이터 = 가구·가전 목록과 구매 연도 + 설비 시공 이력(배관·배선·방수·난방) + 업체별 공종 실적', options: { bold: true } }],
  { x: 0.85, y: 1.95, w: 11.7, h: 0.62, fontSize: 13, valign: 'middle' });

const cols = [
  { x: 0.6, fill: CARD, dark: false, tag: '4-A', name: '주거 보장', one: '쌓인 이력으로 보험사와 주택 보증',
    rows: [['하는 일', '보증 가격의 근거, 수리를 맞는 업체에 한 번에 배정'], ['수익원', '보험사 판매 수수료(위험은 보험사가 짐)'], ['지금 할 일', '보험대리점 제휴로 이사 때 보증금·누수·가재도구 안내, 하자 보수 기한 알림'], ['걸림돌', '세입자 수리는 집주인 몫, 신축은 법정 하자 보수 기간, 제3자 보증의 보험 해당 소지']],
    src: '근거: Frontdoor(보증 매출 약 76%가 갱신, 원가 중심은 수리업체·부품), Porch(집 정보로 보험 가격) 공시' },
  { x: 6.83, fill: DARK, dark: true, tag: '4-B', name: '스마트홈', one: '기기 연동으로 우리 집 데이터의 해상도를 높인다',
    rows: [['하는 일', '가전·설비 상태와 고장 신호를 받아 점검·소모품·수리를 먼저 제안'], ['수익원', '거주 중 관리 매출 증가(정기 관리·소모품·수리), 건설사·관리 앱 제휴'], ['지금 할 일', '기기는 만들지 않고 삼성 SmartThings·LG ThinQ·단지 홈 IoT에 연동, 고객 동의로 상태만 받기'], ['걸림돌', '연동 개방 범위, 생활 패턴 데이터 동의, 직방 홈IoT 인수 뒤 경영 악화 지목(보도)']],
    src: '근거: 제조사 연동·국내 보급 근거는 보강 중(리서치 진행), 직방 사례는 언론 보도' },
];
cols.forEach((c) => {
  const w = 5.9, y = 2.8, h = 3.75, fg = c.dark ? WHITE : INK, sub = c.dark ? 'D0D0D0' : SUB, lab = c.dark ? 'A8B4FF' : ACC;
  rect(c.x, y, w, h, c.fill);
  T([{ text: c.tag + '  ', options: { color: c.dark ? (team ? 'A8B4FF' : '9EC1F0') : ACC } }, { text: c.name, options: { color: fg } }], { x: c.x + 0.25, y: y + 0.2, w: 5.4, h: 0.45, fontSize: 18, bold: true });
  T(c.one, { x: c.x + 0.25, y: y + 0.68, w: 5.4, h: 0.35, fontSize: 12.5, bold: true, color: c.dark ? 'FF8A65' : ORG });
  c.rows.forEach((r, i) => {
    const yy = y + 1.15 + i * 0.6;
    T(r[0], { x: c.x + 0.25, y: yy, w: 1.15, h: 0.55, fontSize: 11, bold: true, color: lab });
    T(r[1], { x: c.x + 1.4, y: yy, w: 4.3, h: 0.55, fontSize: 11, color: sub });
  });
  T(c.src, { x: c.x + 0.25, y: y + 3.4, w: 5.4, h: 0.3, fontSize: 9, color: c.dark ? '9A9A9A' : MUTE });
});
// 두 줄 사이 화살표
s.addShape('leftRightArrow', { x: 6.47, y: 4.35, w: 0.4, h: 0.32, fill: { color: ACC }, line: { color: ACC } });
T('스마트홈 상태 데이터 → 보증 가격·수리 원가 정확도 ↑ / 보증 상품 → 고객이 연동을 켤 이유', { x: 0.6, y: 6.7, w: 12.13, h: 0.3, fontSize: 10.5, bold: true, color: ACC, align: 'center' });
p.writeFile({ fileName: OUT }).then(() => console.log('done'));
