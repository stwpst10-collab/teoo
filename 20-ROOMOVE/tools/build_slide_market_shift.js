// 스토리라인 1번: 주거 공간 산업의 구조적 변화 4가지(팀 장표 톤, 1장)
const pptxgen = require('pptxgenjs');
const OUT = process.argv[2];
const F = 'Malgun Gothic';
const INK = '111111', SUB = '444444', MUTE = '8A8A8A', CARD = 'F2F2F2', BLACK = '161616', WHITE = 'FFFFFF';
const V = ['8E8AFF', '6F6BFA', '5F5BF7', '3B37E8'];
const ACC = '3B37E8', RED = 'D4380D';
const p = new pptxgen(); p.layout = 'LAYOUT_WIDE'; p.theme = { headFontFace: F, bodyFontFace: F };
const s = p.addSlide(); s.background = { color: WHITE };
const T = (t, o) => s.addText(t, Object.assign({ isTextBox: true, fontFace: F, color: INK, margin: 0, valign: 'top' }, o));
const rect = (x, y, w, h, c) => s.addShape('rect', { x, y, w, h, fill: { color: c }, line: { color: c } });

T('Market', { x: 0.6, y: 0.45, w: 8, h: 0.3, fontSize: 13, bold: true });
T('주거 공간 산업은 네 방향으로 바뀌고 있다', { x: 0.6, y: 0.8, w: 12.1, h: 0.6, fontSize: 26, bold: true });
T('기업은 "집을 아는 것"에, 고객은 "더 편한 것"에 돈을 쓰기 시작했다.', { x: 0.6, y: 1.42, w: 12.1, h: 0.35, fontSize: 14, bold: true, color: '333333' });

const cols = [
  { n: '01', t: '집의 데이터화', one: '기기가 집 데이터를 모으고\n서비스를 먼저 제안한다',
    big: '2,000만 명', bigl: 'SmartThings 국내 이용자(등록 기준, 2024.12)',
    pts: ['LG 씽큐 온: AI 홈 허브 정식 출시(2025.10)', 'LG 씽큐 API 외부 개방(2024.12)', '직방: 삼성SDS 홈IoT 인수(2022)'] },
  { n: '02', t: '서비스의 진화', one: '물건 하나를 팔던 회사가\n집 전체를 맡는 회사로',
    big: '1,000만 건', bigl: '미소 누적 주문, 서비스 1,000개 이상(2025.9)',
    pts: ['오늘의집: 커머스 → 시공 → 이사, 직접 부분 시공', 'LG 가전 구독(관리 포함) 매출 연 1조 원 넘음(2023)', '청소연구소 정기 이용 고객 65%'] },
  { n: '03', t: '공간의 가변화', one: '한번 정하면 끝인 집에서\n삶에 따라 바꾸는 집으로',
    big: '27년', bigl: '한국 아파트 평균 사용연수(2026.6 보도)',
    pts: ['삼성물산 개포우성7차 장수명 주택 우수등급 설계(2025.7)', '라멘 구조: 모듈 벽체를 옮겨 구조 변경', '리모델링 인테리어 13.8조원(2025 전망, 2021년 추정)'] },
  { n: '04', t: '로봇 친화 설계', one: '사람만 쓰던 집에서\n로봇과 함께 쓰는 집으로',
    big: '180억 원', bigl: '국토부 로봇 친화 건축 기술개발(2025~2028)',
    pts: ['현대건설 디에이치 대치 에델루이 배송·순찰 로봇', '삼성물산 래미안 리더스원 문 앞 배달 로봇', '로봇 친화 인증 시 용적률·높이 완화 추진'] },
];
const x0 = 0.6, gap = 0.18, w = (12.13 - gap * 3) / 4, y = 2.05, h = 3.9;
cols.forEach((c, i) => {
  const x = x0 + i * (w + gap);
  rect(x, y, w, 0.62, V[i]);
  T([{ text: c.n + '  ', options: { color: 'E0DEFF' } }, { text: c.t }], { x: x + 0.2, y, w: w - 0.3, h: 0.62, fontSize: 15, bold: true, color: WHITE, valign: 'middle' });
  rect(x, y + 0.62, w, h - 0.62, CARD);
  T(c.one, { x: x + 0.2, y: y + 0.8, w: w - 0.35, h: 0.6, fontSize: 12, bold: true });
  T(c.big, { x: x + 0.2, y: y + 1.5, w: w - 0.35, h: 0.55, fontSize: 24, bold: true, color: RED });
  T(c.bigl, { x: x + 0.2, y: y + 2.05, w: w - 0.35, h: 0.45, fontSize: 9.5, color: MUTE });
  T(c.pts.map((t, j) => ({ text: t, options: { bullet: { indent: 10 }, breakLine: j < c.pts.length - 1 } })), { x: x + 0.2, y: y + 2.55, w: w - 0.35, h: 1.3, fontSize: 10, color: SUB, paraSpaceAfter: 3 });
});
rect(0.6, 6.1, 12.13, 0.62, BLACK);
T([{ text: '그래서  ', options: { color: 'BDBDBD' } }, { text: '집이 데이터로 연결되고, 서비스·공간·로봇이 집을 따라 바뀐다. 이 변화가 한꺼번에 몰리는 순간이 이사다.', options: { color: WHITE } }],
  { x: 0.85, y: 6.1, w: 11.7, h: 0.62, fontSize: 13.5, bold: true, valign: 'middle' });
T('자료: 삼성전자 발표(SmartThings), LG전자 발표(씽큐 온·API), 유니콘팩토리(2025.9.30, 미소 회사 발표), 버킷플레이스 감사보고서, LG 구독 언론 보도, 아시아경제(2026.6), 삼성물산 보도자료(2025.7), 한국건설산업연구원 전망 인용 보도(2021), 국토교통부·스마트시티(2025), 언론 보도(현대건설·삼성물산 로봇)',
  { x: 0.6, y: 6.85, w: 12.13, h: 0.35, fontSize: 8.5, color: MUTE, align: 'center' });
p.writeFile({ fileName: OUT }).then(() => console.log('done'));
