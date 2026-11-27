// ROOMOVE 확장 논리 1장: 팀 장표(Figma ROOMOVE_Drafts) 톤에 맞춘 판
const pptxgen = require('pptxgenjs');
const OUT = process.argv[2];
const F = 'Malgun Gothic';
const INK = '111111', SUB = '555555', MUTE = '8A8A8A', CARD = 'F2F2F2', BLACK = '161616', WHITE = 'FFFFFF';
const V = ['D9D7FF', 'B9B6FF', '8E8AFF', '5F5BF7', '3B37E8']; // 팀 장표 연보라→파랑
const ACC = '3B37E8';

const p = new pptxgen();
p.layout = 'LAYOUT_WIDE';
p.theme = { headFontFace: F, bodyFontFace: F };
const s = p.addSlide();
s.background = { color: WHITE };
const T = (text, o) => s.addText(text, Object.assign({ isTextBox: true, fontFace: F, color: INK, margin: 0, valign: 'top' }, o));

T('Expansion', { x: 0.6, y: 0.45, w: 6, h: 0.3, fontSize: 13, bold: true });
T('이사에서 시작해, 집의 생애 전체로', { x: 0.6, y: 0.8, w: 12, h: 0.6, fontSize: 28, bold: true });
T('이사 때 쌓인 집 정보가 다음 서비스를 먼저 제안하게 만든다.', { x: 0.6, y: 1.45, w: 12, h: 0.35, fontSize: 14, bold: true, color: '333333' });

// 5단계 화살표 띠
const st = [
  ['이사·입주', '업종 사이 일정 조율', '업체 수수료'],
  ['입주 직후', '청소·설치·인터넷·수리 연결', '성사 수수료'],
  ['거주 중 관리', '정기 관리·소모품 제안', '정기 이용료·제휴 수수료'],
  ['보장·IoT', '쌓인 이력으로 주택 보증', '보험사 판매 수수료'],
  ['다음 이사·거래', '재이사·임대·매매 지원', '제휴사 이용료'],
];
const x0 = 0.6, W = 12.13, n = st.length, w = W / n, y = 2.15, h = 0.62;
[...st.keys()].reverse().forEach((i) => {
  const x = x0 + i * w;
  s.addShape(i === 0 ? 'homePlate' : 'chevron', { x: i === 0 ? x : x - 0.18, y, w: i === 0 ? w + 0.18 : (i === n - 1 ? w + 0.18 : w + 0.36), h, fill: { color: V[i] }, line: { color: V[i] } });
});
st.forEach((r, i) => {
  const x = x0 + i * w;
  T(r[0], { x: x + 0.1, y, w: w - 0.2, h, fontSize: 15, bold: true, align: 'center', valign: 'middle', color: i >= 3 ? WHITE : INK });
  if (i === 0) T('현재', { x: x + 0.1, y: y - 0.32, w: w - 0.2, h: 0.28, fontSize: 10, bold: true, color: ACC, align: 'center' });
  if (i === 3) T('장기', { x: x + 0.1, y: y - 0.32, w: w - 0.2, h: 0.28, fontSize: 10, bold: true, color: ACC, align: 'center' });
  s.addShape('ellipse', { x: x + w / 2 - 0.07, y: y + 0.9, w: 0.14, h: 0.14, fill: { color: ACC }, line: { color: ACC } });
  T(r[1], { x: x + 0.1, y: y + 1.12, w: w - 0.2, h: 0.32, fontSize: 12, bold: true, align: 'center' });
  T(r[2], { x: x + 0.1, y: y + 1.45, w: w - 0.2, h: 0.3, fontSize: 10.5, color: MUTE, align: 'center' });
});
s.addShape('line', { x: x0, y: y + 0.97, w: W, h: 0, line: { color: 'D0D0D0', width: 1, dashType: 'dash' } });

// 하단: 쌓이는 것(회색) / 그래서(검정)
const by = 4.25, bh = 2.25;
s.addShape('rect', { x: 0.6, y: by, w: 5.9, h: bh, fill: { color: CARD }, line: { color: CARD } });
T('이사 때 쌓이는 집 정보', { x: 0.85, y: by + 0.25, w: 5.4, h: 0.4, fontSize: 16, bold: true });
T('어디서 샀든, 어느 업체든 한 집 기준으로 남는다', { x: 0.85, y: by + 0.68, w: 5.4, h: 0.3, fontSize: 11.5, color: SUB });
['새집 평면·치수', '가구·가전 목록과 구매 연도', '업체별 일정 이행 기록', '고장·수리·교체 이력'].forEach((t, i) => {
  T(t, { x: 0.85 + (i % 2) * 2.7, y: by + 1.15 + Math.floor(i / 2) * 0.42, w: 2.6, h: 0.35, fontSize: 12.5, bold: true });
});

s.addShape('rect', { x: 6.83, y: by, w: 5.9, h: bh, fill: { color: BLACK }, line: { color: BLACK } });
T('그래서', { x: 7.08, y: by + 0.25, w: 5.4, h: 0.3, fontSize: 11.5, bold: true, color: 'BDBDBD' });
T('고객이 늘수록 기록이 쌓이고,\n기록이 쌓일수록 제안이 정확해진다.', { x: 7.08, y: by + 0.58, w: 5.4, h: 0.8, fontSize: 17, bold: true, color: WHITE });
T([{ text: '먼저 확인할 것  ', options: { bold: true, color: 'BDBDBD' } }, { text: '정기 서비스 전환율 · 고객 확보 비용 · 집 정보 등록률', options: { color: WHITE } }],
  { x: 7.08, y: by + 1.6, w: 5.4, h: 0.45, fontSize: 11.5 });

T('해외 사례: 贝壳(중개 → 인테리어·임대, 매출 비중 약 39%), Frontdoor(거래 시점 → 매년 갱신 보증), Porch(검사 소프트웨어 → 보험). 전환율·수수료는 측정 전 가설', { x: 0.6, y: 6.8, w: 12.13, h: 0.3, fontSize: 9, color: MUTE, align: 'center' });

p.writeFile({ fileName: OUT }).then(() => console.log('done'));
