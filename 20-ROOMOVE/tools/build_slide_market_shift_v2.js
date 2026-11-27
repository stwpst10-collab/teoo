// 스토리라인 1번 v2: 왼쪽 구조적 변화 4가지 → 오른쪽 숫자(차트 2개 + 벤치마크), 팀 장표 톤
const pptxgen = require('pptxgenjs');
const OUT = process.argv[2];
const F = 'Malgun Gothic';
const INK = '111111', SUB = '444444', MUTE = '8A8A8A', CARD = 'F2F2F2', BLACK = '161616', WHITE = 'FFFFFF', GRAYBAR = 'BDBDBD';
const ACC = '3B37E8', RED = 'D4380D', LAV = 'E9E8FF';
const p = new pptxgen(); p.layout = 'LAYOUT_WIDE'; p.theme = { headFontFace: F, bodyFontFace: F };
const s = p.addSlide(); s.background = { color: WHITE };
const T = (t, o) => s.addText(t, Object.assign({ isTextBox: true, fontFace: F, color: INK, margin: 0, valign: 'top' }, o));
const rect = (x, y, w, h, c) => s.addShape('rect', { x, y, w, h, fill: { color: c }, line: { color: c } });

T('Market', { x: 0.6, y: 0.4, w: 8, h: 0.3, fontSize: 13, bold: true });
T('이사는 줄어도, 집을 맡기는 시장은 커진다', { x: 0.6, y: 0.72, w: 12.1, h: 0.6, fontSize: 26, bold: true });
T('주거 공간 산업은 네 방향으로 바뀌고 있고, 돈은 "집을 아는 쪽"과 "대신 해 주는 쪽"으로 몰린다.', { x: 0.6, y: 1.32, w: 12.1, h: 0.35, fontSize: 13.5, bold: true, color: '333333' });

// 왼쪽: 구조적 변화
T('시장 내 진행 중인 구조적 변화', { x: 0.6, y: 1.9, w: 3.9, h: 0.3, fontSize: 12, bold: true, color: SUB });
const shifts = [['집의 데이터화', '기기가 먼저 제안한다'], ['서비스의 진화', '물건에서 집 전체로'], ['공간의 가변화', '삶 따라 바꾸는 집'], ['로봇 친화 설계', '로봇과 함께 쓰는 집']];
shifts.forEach((r, i) => {
  const y = 2.3 + i * 1.08;
  rect(0.6, y, 3.75, 0.95, CARD);
  s.addShape('ellipse', { x: 0.8, y: y + 0.22, w: 0.5, h: 0.5, fill: { color: ACC }, line: { color: ACC } });
  T(String(i + 1).padStart(2, '0'), { x: 0.8, y: y + 0.22, w: 0.5, h: 0.5, fontSize: 12, bold: true, color: WHITE, align: 'center', valign: 'middle' });
  T(r[0], { x: 1.45, y: y + 0.15, w: 2.8, h: 0.3, fontSize: 11, color: ACC, bold: true });
  T(r[1], { x: 1.45, y: y + 0.45, w: 2.8, h: 0.4, fontSize: 15, bold: true });
});
s.addShape('chevron', { x: 4.5, y: 3.75, w: 0.35, h: 0.75, fill: { color: '2E9E6B' }, line: { color: '2E9E6B' } });

// 오른쪽 위: 차트 2개
T('숫자로 본 변화', { x: 5.05, y: 1.9, w: 7.6, h: 0.3, fontSize: 12, bold: true, color: SUB });
rect(5.05, 2.3, 3.75, 2.55, WHITE);
const chartBase = { barDir: 'col', showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 10, dataLabelColor: INK, dataLabelFontFace: F,
  catAxisLabelFontFace: F, catAxisLabelFontSize: 10, catAxisLabelColor: SUB, valAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
  showLegend: false, showTitle: true, titleFontFace: F, titleFontSize: 11, titleColor: INK, barGapWidthPct: 60 };
s.addChart(p.charts.BAR, [{ name: '이동자', labels: ['2021', '2022', '2023', '2024', '2025'], values: [721.3, 615.2, 612.9, 628.3, 611.8] }],
  Object.assign({}, chartBase, { x: 5.05, y: 2.3, w: 3.75, h: 2.55, chartColors: [GRAYBAR], dataLabelFormatCode: '0', title: '국내 인구 이동자(만 명)' }));
s.addChart(p.charts.BAR, [{ name: '매출', labels: ['2022', '2023', '2024', '2025'], values: [1864, 2402, 2879, 3216] }],
  Object.assign({}, chartBase, { x: 8.95, y: 2.3, w: 3.78, h: 2.55, chartColors: [ACC], dataLabelFormatCode: '#,##0', title: '오늘의집 매출(억 원)' }));
pill(5.25, 4.85, 3.35, 0.32, CARD); T('5년간 −15%, 2025년 1974년 이후 최저', { x: 5.25, y: 4.85, w: 3.35, h: 0.32, fontSize: 10, bold: true, color: SUB, align: 'center', valign: 'middle' });
pill(9.15, 4.85, 3.38, 0.32, LAV); T('3년간 +73%, 커머스 → 시공 → 이사', { x: 9.15, y: 4.85, w: 3.38, h: 0.32, fontSize: 10, bold: true, color: ACC, align: 'center', valign: 'middle' });
function pill(x, y, w, h, c) { s.addShape('roundRect', { x, y, w, h, fill: { color: c }, line: { color: c }, rectRadius: h / 2 }); }

// 오른쪽 아래: 벤치마크
T('집 전체를 맡는 쪽으로 간 곳들', { x: 5.05, y: 5.3, w: 7.6, h: 0.28, fontSize: 11, bold: true, color: SUB });
const bm = [['미소', '1,000만 건', '누적 주문, 서비스 1,000개+'], ['LG 가전 구독', '2조 원', '2025 연 매출 돌파 전망'], ['SmartThings', '2,000만 명', '국내 이용자(등록 기준)'], ['贝壳', '39%', '인테리어·임대 매출 비중']];
const bw = (7.68 - 0.15 * 3) / 4;
bm.forEach((b, i) => {
  const x = 5.05 + i * (bw + 0.15);
  rect(x, 5.62, bw, 1.0, i === 1 ? BLACK : CARD);
  T(b[0], { x: x + 0.15, y: 5.68, w: bw - 0.25, h: 0.25, fontSize: 9.5, bold: true, color: i === 1 ? 'BDBDBD' : SUB });
  T(b[1], { x: x + 0.15, y: 5.92, w: bw - 0.25, h: 0.38, fontSize: 17, bold: true, color: i === 1 ? 'FF8A65' : RED });
  T(b[2], { x: x + 0.15, y: 6.3, w: bw - 0.25, h: 0.28, fontSize: 8.5, color: i === 1 ? 'D0D0D0' : MUTE });
});
T('자료: 국가데이터처·통계청 국내인구이동통계(2021은 전년 대비 감소분으로 역산), 버킷플레이스 감사보고서(2022~2025, 연결), 유니콘팩토리(2025.9, 미소 회사 발표), LG전자 구독 매출 언론 보도(2025.11, 전망치), 삼성전자 발표(SmartThings, 2024.12), KE Holdings 2025 연간실적. 인과 관계는 당사 해석',
  { x: 0.6, y: 6.85, w: 12.13, h: 0.4, fontSize: 8.5, color: MUTE, align: 'center' });
p.writeFile({ fileName: OUT }).then(() => console.log('done'));
