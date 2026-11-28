// 주거 산업의 흐름 한 장: "무엇에서 → 무엇으로" 4줄 + 근거 숫자 하나씩(팀 장표 톤)
const pptxgen = require('pptxgenjs');
const OUT = process.argv[2];
const F = 'Malgun Gothic';
const INK = '111111', SUB = '555555', MUTE = '8A8A8A', CARD = 'F2F2F2', BLACK = '161616', WHITE = 'FFFFFF';
const ACC = '3B37E8', RED = 'D4380D', LAV = 'EEEDFF';
const p = new pptxgen(); p.layout = 'LAYOUT_WIDE'; p.theme = { headFontFace: F, bodyFontFace: F };
const s = p.addSlide(); s.background = { color: WHITE };
const T = (t, o) => s.addText(t, Object.assign({ isTextBox: true, fontFace: F, color: INK, margin: 0, valign: 'top' }, o));
const rect = (x, y, w, h, c) => s.addShape('rect', { x, y, w, h, fill: { color: c }, line: { color: c } });

T('Market', { x: 0.6, y: 0.4, w: 8, h: 0.3, fontSize: 13, bold: true });
T('주거 산업은 "사는 것"에서 "맡기는 것"으로 흐르고 있다', { x: 0.6, y: 0.72, w: 12.1, h: 0.6, fontSize: 26, bold: true });
T('물건과 거래는 멈췄고, 집을 대신 관리하는 서비스와 데이터로 돈이 옮겨 간다.', { x: 0.6, y: 1.32, w: 12.1, h: 0.35, fontSize: 14, bold: true, color: '333333' });

T('지금까지', { x: 0.6, y: 1.95, w: 3.4, h: 0.28, fontSize: 11, bold: true, color: MUTE });
T('앞으로', { x: 4.75, y: 1.95, w: 3.6, h: 0.28, fontSize: 11, bold: true, color: ACC });
T('근거', { x: 8.75, y: 1.95, w: 3.9, h: 0.28, fontSize: 11, bold: true, color: MUTE });
const rows = [
  ['물건을 산다', '집을 맡긴다', '+2.5% vs +374%', '온라인 가구 거래액 vs 숨고 매출(2021→2025)'],
  ['한 번 사고 끝', '계속 관리받는다', '1.6조 → 2조 원', 'LG 가전 구독 국내 매출(2024→2025)'],
  ['업종마다 따로', '집 전체를 한 곳에서', '+173%', '오늘의집 매출(2021→2025), 커머스 → 시공 → 이사'],
  ['사람이 기억한다', '집 데이터가 기억한다', '2,000만 명', 'SmartThings 국내 이용자(등록 기준, 2024.12)'],
];
rows.forEach((r, i) => {
  const y = 2.3 + i * 0.98, h = 0.82;
  rect(0.6, y, 3.55, h, CARD);
  T(r[0], { x: 0.85, y, w: 3.2, h, fontSize: 17, bold: true, color: SUB, valign: 'middle' });
  s.addShape('rightArrow', { x: 4.22, y: y + 0.24, w: 0.45, h: 0.34, fill: { color: ACC }, line: { color: ACC } });
  rect(4.75, y, 3.75, h, LAV);
  T(r[1], { x: 5.0, y, w: 3.4, h, fontSize: 17, bold: true, color: ACC, valign: 'middle' });
  T(r[2], { x: 8.75, y: y + 0.05, w: 3.95, h: 0.42, fontSize: 18, bold: true, color: RED });
  T(r[3], { x: 8.75, y: y + 0.47, w: 3.95, h: 0.32, fontSize: 9.5, color: MUTE });
});
rect(0.6, 6.25, 12.13, 0.55, BLACK);
T([{ text: '그래서  ', options: { color: 'BDBDBD' } }, { text: '이 흐름이 한꺼번에 몰리는 순간이 이사다. 이사에서 집을 맡고, 그 집을 계속 관리한다.', options: { color: WHITE } }],
  { x: 0.85, y: 6.25, w: 11.7, h: 0.55, fontSize: 13.5, bold: true, valign: 'middle' });
T('자료: 통계청 온라인쇼핑동향(가구, 2025 잠정), 숨고·버킷플레이스 감사보고서(DART, 오늘의집 별도), LG전자 실적 발표, 삼성전자 발표. 시계열 상세는 시계열 장표 참고, 흐름 해석은 당사 판단',
  { x: 0.6, y: 6.92, w: 12.13, h: 0.3, fontSize: 8.5, color: MUTE, align: 'center' });
p.writeFile({ fileName: OUT }).then(() => console.log('done'));
