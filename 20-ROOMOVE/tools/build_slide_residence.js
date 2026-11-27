// 거주 중 관리(반복 매출) 상세 장표 1장: 팀 장표 톤
const pptxgen = require('pptxgenjs');
const OUT = process.argv[2];
const F = 'Malgun Gothic';
const INK = '111111', SUB = '444444', MUTE = '8A8A8A', CARD = 'F2F2F2', BLACK = '161616', WHITE = 'FFFFFF';
const V = ['D9D7FF', 'B9B6FF', '8E8AFF', '5F5BF7', '3B37E8'];
const ACC = '3B37E8', RED = 'D4380D';
const p = new pptxgen(); p.layout = 'LAYOUT_WIDE'; p.theme = { headFontFace: F, bodyFontFace: F };
const s = p.addSlide(); s.background = { color: WHITE };
const T = (t, o) => s.addText(t, Object.assign({ isTextBox: true, fontFace: F, color: INK, margin: 0, valign: 'top' }, o));
const rect = (x, y, w, h, c) => s.addShape('rect', { x, y, w, h, fill: { color: c }, line: { color: c } });
const pill = (x, y, w, h, c) => s.addShape('roundRect', { x, y, w, h, fill: { color: c }, line: { color: c }, rectRadius: h / 2 });

T('Expansion · 거주 중 관리', { x: 0.6, y: 0.45, w: 8, h: 0.3, fontSize: 13, bold: true });
T('반복 매출은 거주 중 관리에서 나온다', { x: 0.6, y: 0.8, w: 12.1, h: 0.6, fontSize: 26, bold: true });
T('이사 때 받은 우리 집 데이터로, 필요한 관리를 고객보다 먼저 제안한다.', { x: 0.6, y: 1.42, w: 12.1, h: 0.35, fontSize: 14, bold: true, color: '333333' });

// 왼쪽: 데이터 → 제안 시점 → 서비스
const flow = [
  ['우리 집 데이터', '가전·설비 모델·구매 연도, 인테리어 시공 이력(배관·배선·방수·난방), 지난 청소·수리 이력'],
  ['제안 시점', '필터 교체 주기, 에어컨 청소 시즌, 하자 보수 기간 끝나기 전 점검'],
  ['연결하는 서비스', '정기 청소 · 소모품 정기 구매 · 수리는 그 공종을 잘해 온 업체에 배정'],
];
flow.forEach((f, i) => {
  const y = 2.15 + i * 1.25;
  pill(0.6, y, 2.1, 0.5, V[i + 2]);
  T(f[0], { x: 0.6, y, w: 2.1, h: 0.5, fontSize: 13, bold: true, align: 'center', valign: 'middle', color: WHITE });
  T(f[1], { x: 0.6, y: y + 0.58, w: 5.2, h: 0.55, fontSize: 12, color: SUB });
  if (i < 2) T('↓', { x: 1.45, y: y + 1.0, w: 0.4, h: 0.25, fontSize: 12, bold: true, color: ACC, align: 'center' });
});

// 가운데: 고객 1명 연 수수료
rect(6.05, 2.15, 3.35, 3.65, CARD);
T('고객 1명 연 수수료(당사 추정)', { x: 6.25, y: 2.3, w: 3.0, h: 0.3, fontSize: 11.5, bold: true, color: SUB });
const rev = [['정기 청소', '2주 1회', '17.4만원', INK], ['에어컨 분해 청소', '2대·연 1회', '1.6만~2.6만원', INK], ['소모품 정기 구매', '제휴 3% 안팎', '측정 전', MUTE], ['수리·교체 연결', '건별', '보조 수익', MUTE]];
rev.forEach((r, i) => {
  const y = 2.75 + i * 0.75;
  T(r[0], { x: 6.25, y, w: 1.75, h: 0.3, fontSize: 12.5, bold: true });
  T(r[1], { x: 6.25, y: y + 0.3, w: 1.75, h: 0.28, fontSize: 10, color: MUTE });
  T(r[2], { x: 7.85, y, w: 1.4, h: 0.55, fontSize: i === 0 ? 16 : 12.5, bold: true, color: i === 0 ? RED : r[3], align: 'right', valign: 'middle' });
});

// 오른쪽: 왜 반복 매출의 키인가
rect(9.65, 2.15, 3.08, 3.65, BLACK);
T('왜 반복 매출의 키인가', { x: 9.85, y: 2.3, w: 2.8, h: 0.3, fontSize: 11.5, bold: true, color: 'BDBDBD' });
T('3.5~4.4배', { x: 9.85, y: 2.7, w: 2.8, h: 0.6, fontSize: 28, bold: true, color: 'FF8A65' });
T('거주 중 서비스가 붙으면 고객 1명 5년 수수료가 22만~33만원에서 98만~113만원으로', { x: 9.85, y: 3.35, w: 2.75, h: 0.85, fontSize: 11, color: 'E0E0E0' });
T('3.7년 · 11.5년', { x: 9.85, y: 4.3, w: 2.8, h: 0.45, fontSize: 18, bold: true, color: WHITE });
T('평균 거주기간(임차·자가). 다음 이사보다 사는 기간이 훨씬 길다', { x: 9.85, y: 4.8, w: 2.75, h: 0.8, fontSize: 11, color: 'E0E0E0' });

// 하단: 차별점과 위험
rect(0.6, 6.0, 12.13, 0.72, CARD);
T([{ text: '차별점  ', options: { bold: true, color: ACC } }, { text: '인테리어·이사를 조율하며 벽 속 설비 이력까지 남는 곳은 ROOMOVE뿐이다(가설). 보증 사업의 이익을 가르는 수리 원가를 낮추는 정보가 된다.', options: { bold: true } },
   { text: '', options: { color: SUB } }],
  { x: 0.85, y: 6.0, w: 11.7, h: 0.72, fontSize: 11.5, valign: 'middle' });
T('자료: 미소 정기 청소 요금·숨고·LG전자 공개 가격(2026.11, 업체 자료), 공정위 부품 보유기간, 2025년도 주거실태조사, 청소연구소 발표(정기 이용 65%·재구매 88%), Frontdoor 2025 실적 발표(매출총이익률 55%, 원가는 수리업체·부품 비용), 당사 추정. 전환율·유지율은 측정 전 가설',
  { x: 0.6, y: 6.85, w: 12.13, h: 0.3, fontSize: 9, color: MUTE, align: 'center' });
p.writeFile({ fileName: OUT }).then(() => console.log('done'));
