// 경쟁사 포지셔닝 맵(사분면): 보고서 v2 2장 기준, 팀 장표 톤
const pptxgen = require('pptxgenjs');
const OUT = process.argv[2];
const F = 'Malgun Gothic';
const INK = '111111', SUB = '444444', MUTE = '8A8A8A', CARD = 'F2F2F2', BLACK = '161616', WHITE = 'FFFFFF';
const ACC = '3B37E8', RED = 'D4380D', LAV = 'E9E8FF', BUB = '8E8AFF';
const p = new pptxgen(); p.layout = 'LAYOUT_WIDE'; p.theme = { headFontFace: F, bodyFontFace: F };
const s = p.addSlide(); s.background = { color: WHITE };
const T = (t, o) => s.addText(t, Object.assign({ isTextBox: true, fontFace: F, color: INK, margin: 0, valign: 'top' }, o));
const rect = (x, y, w, h, c, line) => s.addShape('rect', { x, y, w, h, fill: { color: c }, line: { color: line || c } });

T('Competition', { x: 0.6, y: 0.4, w: 8, h: 0.3, fontSize: 13, bold: true });
T('견적 연결은 빼곡하고, 업종 사이 일정을 맡는 자리는 비어 있다', { x: 0.6, y: 0.72, w: 12.1, h: 0.6, fontSize: 24, bold: true });
T('여러 업종을 다루는 곳은 연결에 머물고, 일정을 직접 관리하는 곳은 자기 상품·자기 시공 안에 머문다.', { x: 0.6, y: 1.3, w: 12.1, h: 0.35, fontSize: 13.5, bold: true, color: '333333' });

// 사분면 영역
const X0 = 1.05, Y0 = 1.95, W = 7.4, H = 4.55; // plot box
const qx = X0 + W / 2, qy = Y0 + H / 2;
rect(X0, Y0, W / 2, H / 2, 'FAFAFA'); rect(qx, Y0, W / 2, H / 2, LAV); rect(X0, qy, W / 2, H / 2, 'FAFAFA'); rect(qx, qy, W / 2, H / 2, CARD);
s.addShape('line', { x: X0, y: qy, w: W, h: 0, line: { color: 'C8C8C8', width: 1 } });
s.addShape('line', { x: qx, y: Y0, w: 0, h: H, line: { color: 'C8C8C8', width: 1 } });
const q = [[X0 + 0.12, Y0 + 0.1, '자기 일만 직접 관리', '자사 상품·자사 시공 안의 일정'], [qx + 0.12, Y0 + 0.1, '여러 업종 일정을 직접 맞춤', '공개 자료로 확인된 곳 없음'],
           [X0 + 0.12, Y0 + H - 0.55, '한 업종 연결', '견적·예약 중개'], [qx + W / 2 - 3.52, Y0 + H - 0.55, '여러 업종 연결', '이미 많음']];
q.forEach((r, i) => {
  const al = i === 3 ? 'right' : 'left';
  T(r[2], { x: r[0], y: r[1], w: 3.4, h: 0.25, fontSize: 10.5, bold: true, color: i === 1 ? ACC : SUB, align: al });
  T(r[3], { x: r[0], y: r[1] + 0.24, w: 3.4, h: 0.22, fontSize: 9, color: MUTE, align: al });
});
// 축
s.addShape('line', { x: X0, y: Y0 + H + 0.08, w: W, h: 0, line: { color: SUB, width: 1.25, endArrowType: 'triangle' } });
T('맡는 범위  한 업종 → 여러 업종·주거 단계 전체', { x: X0, y: Y0 + H + 0.14, w: W, h: 0.28, fontSize: 10.5, bold: true, color: SUB, align: 'right' });
s.addShape('line', { x: X0 - 0.1, y: Y0, w: 0, h: H, line: { color: SUB, width: 1.25, beginArrowType: 'triangle' } });
T('일정 책임  견적 연결 → 업종 사이 직접 조율', { x: 0.25, y: Y0, w: 0.35, h: H, fontSize: 10.5, bold: true, color: SUB, vert: 'vert270', align: 'center', valign: 'middle' });

// 회사 점: [이름, X(0~10), Y(0~10), 2025 매출(억 원) 또는 null]
const MAXR = 3216;
const pts = [
  ['헬로입주', 2.4, 1.9, null], ['집닥', 1.6, 3.1, null], ['영구크린', 1.4, 6.7, 205], ['삼성 인테리어핏·IKEA 주방', 2.2, 8.1, null],
  ['짐싸·이사타임', 4.4, 2.7, null], ['아정당', 5.0, 4.0, 2196], ['직방', 6.2, 1.6, 922], ['숨고', 7.1, 2.3, 720],
  ['홈닉', 6.4, 4.3, null], ['미소', 8.9, 2.9, null], ['오늘의집', 8.1, 6.0, 3216],
];
const px = (x) => X0 + (x / 10) * W, py = (y) => Y0 + H - (y / 10) * H;
pts.forEach(([n, x, y, rev]) => {
  const d = rev ? 0.28 + 0.62 * Math.sqrt(rev / MAXR) : 0.26;
  s.addShape('ellipse', { x: px(x) - d / 2, y: py(y) - d / 2, w: d, h: d, fill: rev ? { color: BUB, transparency: 25 } : { color: WHITE }, line: { color: BUB, width: 1.5 } });
  T(n, { x: px(x) - 1.2, y: py(y) + d / 2 + 0.02, w: 2.4, h: 0.24, fontSize: 9.5, bold: true, align: 'center' });
});
// ROOMOVE 목표 위치
const rx = px(8.7), ry = py(8.6);
s.addShape('star5', { x: rx - 0.24, y: ry - 0.24, w: 0.48, h: 0.48, fill: { color: RED }, line: { color: RED } });
T('ROOMOVE(목표)', { x: rx - 1.3, y: ry + 0.26, w: 2.6, h: 0.26, fontSize: 11, bold: true, color: RED, align: 'center' });

// 오른쪽: 범례·시사점
const RX = 8.85, RW = 3.88;
T('읽는 법', { x: RX, y: 1.95, w: RW, h: 0.28, fontSize: 11, bold: true, color: SUB });
s.addShape('ellipse', { x: RX, y: 2.3, w: 0.32, h: 0.32, fill: { color: BUB, transparency: 25 }, line: { color: BUB, width: 1.5 } });
T('원 크기 = 2025년 매출(감사보고서)', { x: RX + 0.42, y: 2.3, w: RW - 0.42, h: 0.32, fontSize: 10, valign: 'middle' });
s.addShape('ellipse', { x: RX + 0.03, y: 2.72, w: 0.26, h: 0.26, fill: { color: WHITE }, line: { color: BUB, width: 1.5 } });
T('빈 원 = 매출 비공개 또는 확인 못 함', { x: RX + 0.42, y: 2.7, w: RW - 0.42, h: 0.3, fontSize: 10, valign: 'middle' });
T('오늘의집 3,216억 · 아정당 2,196억 · 직방 922억 · 숨고 720억 · 영구크린 205억', { x: RX, y: 3.08, w: RW, h: 0.45, fontSize: 9, color: MUTE });
rect(RX, 3.6, RW, 2.9, BLACK);
T('시사점', { x: RX + 0.2, y: 3.75, w: RW - 0.4, h: 0.28, fontSize: 11, bold: true, color: 'BDBDBD' });
T([{ text: '여러 업종 견적·신청 연결은 이미 빼곡하다(오른쪽 아래)', options: { bullet: { indent: 12 }, breakLine: true } },
   { text: '일정을 직접 관리하는 곳은 자기 상품·자기 시공 안에서만 한다(왼쪽 위)', options: { bullet: { indent: 12 }, breakLine: true } },
   { text: '가장 가까운 위협은 오늘의집. 지정일 배송·부분 시공을 직접 하며 위로 올라오는 중', options: { bullet: { indent: 12 }, breakLine: true } },
   { text: '어디서 샀든 여러 업종 일정을 맞추고 다시 조정하는 자리(오른쪽 위)가 ROOMOVE의 목표', options: { bullet: { indent: 12 } } }],
  { x: RX + 0.2, y: 4.1, w: RW - 0.4, h: 2.35, fontSize: 10.5, color: WHITE, paraSpaceAfter: 6 });
T('자료: 보고서 v2 2장(표 1·2·7·10-1), 각 사 감사보고서(DART, 2025)·앱 소개·언론 보도(2026.11), 당사 정리. 점의 위치는 공개 자료 기준 정성 판단이며, 상담 단계의 비공식 조율은 직접 확인 필요',
  { x: 0.6, y: 6.95, w: 12.13, h: 0.3, fontSize: 8.5, color: MUTE, align: 'center' });
p.writeFile({ fileName: OUT }).then(() => console.log('done'));
