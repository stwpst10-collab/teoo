// 주거 산업 시계열 장표 2장(팀 톤): ① 전체 방향(2021=100 지수 + 변화율 순위) ② 하위 항목 7개 추이
// 숫자 출처: 리서치/확장/주거 산업 시계열 2026-11-27.md
const pptxgen = require('pptxgenjs');
const OUT = process.argv[2];
const F = 'Malgun Gothic';
const INK = '111111', SUB = '444444', MUTE = '8A8A8A', CARD = 'F2F2F2', BLACK = '161616', WHITE = 'FFFFFF';
const ACC = '3B37E8', RED = 'D4380D', G1 = '9E9E9E', G2 = 'BDBDBD', G3 = '6B6B6B';
const p = new pptxgen(); p.layout = 'LAYOUT_WIDE'; p.theme = { headFontFace: F, bodyFontFace: F };
let s;
const T = (t, o) => s.addText(t, Object.assign({ isTextBox: true, fontFace: F, color: INK, margin: 0, valign: 'top' }, o));
const rect = (x, y, w, h, c) => s.addShape('rect', { x, y, w, h, fill: { color: c }, line: { color: c } });
function head(label, title, sub) {
  s = p.addSlide(); s.background = { color: WHITE };
  T(label, { x: 0.6, y: 0.4, w: 8, h: 0.3, fontSize: 13, bold: true });
  T(title, { x: 0.6, y: 0.72, w: 12.1, h: 0.6, fontSize: 24, bold: true });
  T(sub, { x: 0.6, y: 1.3, w: 12.1, h: 0.35, fontSize: 13.5, bold: true, color: '333333' });
}
const axis = { catAxisLabelFontFace: F, catAxisLabelFontSize: 10, catAxisLabelColor: SUB, valAxisLabelFontFace: F, valAxisLabelFontSize: 9, valAxisLabelColor: MUTE, valGridLine: { color: 'E6E6E6', size: 0.5 }, catGridLine: { style: 'none' } };

// ① 전체 방향
head('Market · 시계열', '물건을 사는 시장은 멈췄고, 집을 맡기는 서비스는 커졌다', '2021년을 100으로 맞춰 보면 이사·가구 구매는 제자리이거나 줄었고, 플랫폼·홈서비스는 2~5배가 됐다.');
const yrs = ['2021', '2022', '2023', '2024', '2025'];
const idx = (a) => a.map((v) => Math.round((v / a[0]) * 1000) / 10);
const series = [
  { name: '숨고 매출', values: idx([151.8, 313.8, 458.6, 624.2, 719.8]) },
  { name: '오늘의집 매출', values: idx([1176.3, 1864.2, 2355.2, 2879.5, 3215.7]) },
  { name: '온라인 가구 거래액', values: idx([53976, 51950, 53521, 56670, 55346]) },
  { name: '인구 이동자 수', values: idx([721.3, 615.2, 612.9, 628.3, 611.8]) },
  { name: '한샘 매출', values: idx([22314, 20009, 19669, 19084, 17445]) },
];
s.addChart(p.charts.LINE, series.map((x) => ({ name: x.name, labels: yrs, values: x.values })), Object.assign({
  x: 0.6, y: 1.85, w: 7.3, h: 4.7, chartColors: [RED, ACC, G1, G2, G3], lineSize: 2.5, lineDataSymbol: 'circle', lineDataSymbolSize: 6,
  showLegend: true, legendPos: 'b', legendFontFace: F, legendFontSize: 10, legendColor: SUB, valAxisMinVal: 0,
  showTitle: true, title: '2021 = 100 지수', titleFontFace: F, titleFontSize: 11, titleColor: INK }, axis));
// 오른쪽: 2021→2025 변화율 순위
T('2021 → 2025 변화', { x: 8.25, y: 1.85, w: 4.4, h: 0.3, fontSize: 12, bold: true, color: SUB });
const ch = series.map((x) => [x.name, Math.round((x.values[4] - 100) * 10) / 10]).sort((a, b) => b[1] - a[1]);
const maxAbs = Math.max(...ch.map((c) => Math.abs(c[1])));
ch.forEach(([n, v], i) => {
  const y = 2.3 + i * 0.62, mid = 10.15, scale = 2.3 / maxAbs, len = Math.max(0.04, Math.abs(v) * scale);
  T(n, { x: 8.25, y, w: 1.85, h: 0.45, fontSize: 10.5, bold: true, valign: 'middle' });
  const col = v > 10 ? (i === 0 ? RED : ACC) : G1;
  if (v >= 0) rect(mid, y + 0.1, len, 0.26, col); else rect(mid - len, y + 0.1, len, 0.26, col);
  T((v > 0 ? '+' : '') + v + '%', { x: v >= 0 ? mid + len + 0.05 : mid - len - 0.85, y, w: 0.8, h: 0.45, fontSize: 10.5, bold: true, valign: 'middle', align: v >= 0 ? 'left' : 'right', color: v > 10 ? col : SUB });
});
s.addShape('line', { x: 10.15, y: 2.25, w: 0, h: 3.2, line: { color: 'C8C8C8', width: 1 } });
rect(8.25, 5.65, 4.48, 0.9, BLACK);
T([{ text: '+ LG 가전 구독  ', options: { color: 'BDBDBD' } }, { text: '국내 2024년 1.6조 원(+50%) → 2025년 2조 원 돌파', options: { color: WHITE } }], { x: 8.45, y: 5.7, w: 4.15, h: 0.8, fontSize: 11, bold: true, valign: 'middle' });
T('자료: 숨고·버킷플레이스·한샘 감사보고서(DART, 오늘의집은 별도 기준, 한샘 2025 잠정), 통계청 온라인쇼핑동향(가구, 2025 잠정), 통계청·국가데이터처 국내인구이동통계, LG전자 실적 발표(구독). 당사 정리',
  { x: 0.6, y: 6.85, w: 12.13, h: 0.35, fontSize: 8.5, color: MUTE, align: 'center' });

// ② 하위 항목 7개
head('Market · 하위 항목', '하위 항목별로 보면, 성장은 "맡기고 관리받는 쪽"에 몰린다', '거래·이사·가구 구매는 정체하거나 줄고, 홈서비스·구독·플랫폼은 꾸준히 커진다.');
const cards = [
  { k: 'A 거래·중개', n: '다방 매출(억 원)', l: ['20', '21', '22', '23', '24', '25'], v: [274.6, 246.0, 230.7, 209.4, 209.7, 206.5], dir: '중개 플랫폼 매출 감소', up: false },
  { k: 'B 이사·입주', n: '인구 이동자(만 명)', l: ['18', '19', '20', '21', '22', '23', '24', '25'], v: [729.7, 710.4, 773.5, 721.3, 615.2, 612.9, 628.3, 611.8], dir: '이사 물량은 줄어드는 추세', up: false },
  { k: 'C 가구·인테리어', n: '한샘 매출(억 원)', l: ['18', '19', '20', '21', '22', '23', '24', '25'], v: [19285, 17023, 20675, 22314, 20009, 19669, 19084, 17445], dir: '전통 강자는 2021년 정점 뒤 감소', up: false },
  { k: 'D 가구 구매', n: '온라인 가구 거래액(조 원)', l: ['18', '19', '20', '21', '22', '23', '24', '25'], v: [3.13, 3.54, 4.90, 5.40, 5.20, 5.35, 5.67, 5.53], dir: '2021년부터 5조 원대에서 정체', up: false },
  { k: 'C′ 인테리어 플랫폼', n: '오늘의집 매출(억 원)', l: ['18', '19', '20', '21', '22', '23', '24', '25'], v: [72.1, 242.7, 759.2, 1176.3, 1864.2, 2355.2, 2879.5, 3215.7], dir: '커머스 → 시공 → 이사로 계속 성장', up: true },
  { k: 'E 홈서비스 연결', n: '숨고 매출(억 원)', l: ['21', '22', '23', '24', '25'], v: [151.8, 313.8, 458.6, 624.2, 719.8], dir: '4년간 4.7배, 2023년부터 흑자', up: true },
  { k: 'F 구독·관리', n: 'LG 가전 구독 국내(조 원)', l: ['23', '24', '25'], v: [1.07, 1.6, 2.0], dir: '사는 대신 쓰고 관리받는 흐름', up: true },
];
const cw = (12.13 - 0.15 * 3) / 4, chh = 2.25;
cards.forEach((c, i) => {
  const col = i % 4, row = Math.floor(i / 4), x = 0.6 + col * (cw + 0.15), y = 1.85 + row * (chh + 0.15);
  rect(x, y, cw, chh, c.up ? 'EEEDFF' : CARD);
  T(c.k, { x: x + 0.15, y: y + 0.1, w: cw - 0.3, h: 0.25, fontSize: 10.5, bold: true, color: c.up ? ACC : SUB });
  T(c.n, { x: x + 0.15, y: y + 0.36, w: cw - 0.3, h: 0.22, fontSize: 9, color: MUTE });
  s.addChart(p.charts.BAR, [{ name: c.n, labels: c.l, values: c.v }], {
    x: x + 0.05, y: y + 0.55, w: cw - 0.1, h: 1.3, barDir: 'col', chartColors: [c.up ? ACC : G1], barGapWidthPct: 40,
    catAxisLabelFontFace: F, catAxisLabelFontSize: 8, catAxisLabelColor: MUTE, valAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { style: 'none' }, showLegend: false });
  T(c.dir, { x: x + 0.15, y: y + 1.87, w: cw - 0.3, h: 0.32, fontSize: 10.5, bold: true, color: c.up ? ACC : INK });
});
// 8번째 칸: 스마트홈(시계열 없음)
{ const x = 0.6 + 3 * (cw + 0.15), y = 1.85 + (chh + 0.15);
  rect(x, y, cw, chh, BLACK);
  T('G 스마트홈·IoT', { x: x + 0.15, y: y + 0.1, w: cw - 0.3, h: 0.25, fontSize: 10.5, bold: true, color: 'A8B4FF' });
  T('2,000만 명', { x: x + 0.15, y: y + 0.5, w: cw - 0.3, h: 0.5, fontSize: 24, bold: true, color: 'FF8A65' });
  T('SmartThings 국내 이용자(등록 기준, 2024.12)', { x: x + 0.15, y: y + 1.02, w: cw - 0.3, h: 0.4, fontSize: 9.5, color: 'D0D0D0' });
  T('시장 규모 추정은 기관마다 범위가 달라 시계열로 잇지 않음', { x: x + 0.15, y: y + 1.5, w: cw - 0.3, h: 0.6, fontSize: 9.5, color: '9A9A9A' }); }
T('자료: 각 사 감사보고서·잠정실적(DART), 통계청 온라인쇼핑동향(2025 잠정), 통계청·국가데이터처 국내인구이동통계(2018은 역산), LG전자 실적 발표(2023은 2024 증가율로 역산), 삼성전자 발표. 다방=스테이션3, 오늘의집=별도 기준',
  { x: 0.6, y: 6.85, w: 12.13, h: 0.35, fontSize: 8.5, color: MUTE, align: 'center' });
p.writeFile({ fileName: OUT }).then(() => console.log('done'));
