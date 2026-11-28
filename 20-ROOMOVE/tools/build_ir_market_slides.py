"""팀 IR 덱(ROOMOVE IR Deck V1) 톤으로 산업 장표 2장을 만들어 덱 사본의 3·4쪽 뒤에 넣는다.
- 덱의 CONTENT 레이아웃(라벨·제목·부제·출처 자리표시자)을 그대로 쓰고, 색은 덱 테마 값을 쓴다.
사용: python build_ir_market_slides.py <원본 덱.pptx> <출력.pptx>"""
import sys, copy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC, OUT = sys.argv[1], sys.argv[2]
INK = RGBColor(0x11, 0x11, 0x13); SUB = RGBColor(0x6B, 0x6B, 0x70); LIGHT = RGBColor(0xF3, 0xF3, 0xF5)
ACC = RGBColor(0x3D, 0x38, 0xF5); LAV = RGBColor(0xE9, 0xE8, 0xFF); LINE = RGBColor(0xD4, 0xD4, 0xD8)
MUTE = RGBColor(0x9A, 0x9A, 0xA1); WHITE = RGBColor(0xFF, 0xFF, 0xFF); DARK = RGBColor(0x11, 0x11, 0x13)
F = 'Malgun Gothic'

prs = Presentation(SRC)
REF = prs.slides[6]  # OVERVIEW 장표(7쪽)
layout = REF.slide_layout  # CONTENT 레이아웃

def text(sl, x, y, w, h, t, size=11, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0; tf.word_wrap = True; tf.vertical_anchor = anchor
    for i, line in enumerate(t.split('\n')):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); para.alignment = align
        r = para.add_run(); r.text = line; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; r.font.name = F
    return tb

def box(sl, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE, line=None):
    s = sl.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line: s.line.color.rgb = line; s.line.width = Pt(0.75)
    else: s.line.fill.background()
    s.shadow.inherit = False
    return s

def noshadow(shape):
    from pptx.oxml.ns import qn
    spPr = shape._element.spPr
    if spPr.find(qn('a:effectLst')) is None: spPr.append(spPr.makeelement(qn('a:effectLst'), {}))

def fill_ph(sl, label, title, sub, src):
    for ph in sl.placeholders:
        i = ph.placeholder_format.idx
        t = {1: label, 0: title, 2: sub, 3: src}.get(i)
        if t is None: continue
        ph.text_frame.text = t
        # 7쪽(OVERVIEW) 같은 자리의 문단·글자 서식을 그대로 복사(들여쓰기·크기·자간)
        refph = [r for r in REF.placeholders if r.placeholder_format.idx == i][0]
        rp = refph.text_frame.paragraphs[0]; para = ph.text_frame.paragraphs[0]
        if rp._p.pPr is not None:
            if para._p.pPr is not None: para._p.remove(para._p.pPr)
            para._p.insert(0, copy.deepcopy(rp._p.pPr))
        if rp.runs and para.runs and rp.runs[0]._r.rPr is not None:
            r = para.runs[0]._r
            if r.rPr is not None: r.remove(r.rPr)
            r.insert(0, copy.deepcopy(rp.runs[0]._r.rPr))

# ── 장표 A: 산업 구조 변화 → 그래서 이사에서 시작 ──
a = prs.slides.add_slide(layout)
fill_ph(a, 'MARKET', '주거 공간 산업 내 구조적인 변화', '집이 데이터로 연결되고, 사는 것에서 맡기는 것으로 바뀌는 흐름. 그 변화가 가장 먼저 몰리는 순간이 이사',
        '출처: 삼성전자 발표(SmartThings 국내 이용자, 등록 기준 2024.12), 브레이브모바일(숨고) 감사보고서(2021·2025), 국가데이터처 「2025년 인구주택총조사」, 국토교통부·스마트시티(2025). 오른쪽은 당사 판단')
text(a, 0.56, 1.66, 3.6, 0.22, '시장 내 진행 중인 구조적 변화', 10, True, SUB)
shifts = [('01', '집의 데이터화', '기기가 먼저 제안하는 집', 'SmartThings 국내 2,000만 명'),
          ('02', '서비스의 진화', '사는 것에서 맡기는 것으로', '숨고 매출 4.7배(2021→2025)'),
          ('03', '공간의 가변화', '삶에 맞춰 바꾸는 집', '지은 지 30년 넘은 주택 30.6%'),
          ('04', '로봇·스마트홈 설계', '로봇과 함께 쓰는 집', '국토부 로봇 친화 건축 180억 원')]
for i, (n, k, v, ev) in enumerate(shifts):
    y = 1.95 + i * 0.73
    box(a, 0.56, y, 3.55, 0.64, LIGHT)
    c = box(a, 0.7, y + 0.14, 0.36, 0.36, ACC, MSO_SHAPE.OVAL)
    text(a, 0.7, y + 0.14, 0.36, 0.36, n, 9, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(a, 1.18, y + 0.08, 2.85, 0.2, k, 9, True, ACC)
    text(a, 1.18, y + 0.27, 2.85, 0.24, v, 12, True, INK)
    text(a, 1.18, y + 0.47, 2.85, 0.16, ev, 7.5, False, MUTE)
chev = box(a, 4.2, 2.95, 0.28, 0.7, ACC, MSO_SHAPE.CHEVRON)
text(a, 4.62, 1.66, 4.8, 0.22, '그래서, 이사에서 시작', 10, True, SUB)
rows = [('고객을 만나는 시점', '필요할 때 검색', '이사 때 먼저 만남', '고민과 결정이 한꺼번에 몰리는 순간'),
        ('집 데이터', '서비스마다 흩어짐', '한 집 기준 첫 기록', '평면·가전·시공 이력이 처음 생기는 순간'),
        ('돈을 버는 방식', '한 번 거래하고 끝', '거주 중 계속 관리', '맡기는 서비스가 커지는 쪽으로')]
for i, (k, frm, to, why) in enumerate(rows):
    y = 1.95 + i * 0.97
    box(a, 4.62, y, 1.15, 0.86, DARK)
    text(a, 4.7, y, 1.0, 0.86, k, 10, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    box(a, 5.88, y + 0.08, 1.45, 0.4, LIGHT, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(a, 5.88, y + 0.08, 1.45, 0.4, frm, 10, True, SUB, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(a, 7.36, y + 0.08, 0.3, 0.4, '→', 12, True, ACC, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    box(a, 7.7, y + 0.08, 1.73, 0.4, ACC, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(a, 7.7, y + 0.08, 1.73, 0.4, to, 10, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(a, 5.9, y + 0.56, 3.5, 0.24, '· ' + why, 9, False, SUB)

# ── 장표 B: 기업은 집을 아는 것, 고객은 더 편한 것 ──
b = prs.slides.add_slide(layout)
fill_ph(b, 'MARKET', '기업은 ‘집을 아는 것’, 고객은 ‘더 편한 것’을 찾는 중', '모두 집의 한 조각씩만 아는 상태. 여러 브랜드·여러 업체에 걸친 우리 집 전체를 아는 곳은 아직 없음',
        '출처: 삼성물산 보도자료(2026.11), 삼성전자 발표(2024.12), LG전자 실적 발표(2026.1), 직방·버킷플레이스·브레이브모바일 감사보고서(DART), 유니콘팩토리(2025.9, 미소 회사 발표). 분류는 당사 정리')
text(b, 0.56, 1.66, 5.9, 0.22, '집의 어느 조각을 아는가', 10, True, SUB)
players = [('구조·입주일', '건설사 앱', '홈닉: 입주 전 인테리어·청소 신청, 건설사 21곳 협약'),
           ('집 안의 기기', '가전 제조사', 'SmartThings 국내 2,000만 명, LG 씽큐 외부 개방'),
           ('거래·위치', '부동산 플랫폼', '직방: 매물에서 거주 중 등기 알림·다음 이사까지'),
           ('취향·구매', '인테리어 커머스', '오늘의집: 커머스 → 시공 → 이사'),
           ('집안일', '홈서비스 연결', '미소·숨고: 청소·이사·설치·수리 연결')]
for i, (k, who, what) in enumerate(players):
    y = 1.95 + i * 0.55
    box(b, 0.56, y, 1.3, 0.46, LAV, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(b, 0.56, y, 1.3, 0.46, k, 10, True, ACC, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(b, 2.0, y + 0.02, 1.2, 0.42, who, 10, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    text(b, 3.2, y + 0.02, 3.35, 0.42, what, 9, False, SUB, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    ln = b.shapes.add_connector(1, Inches(0.56), Inches(y + 0.51), Inches(6.55), Inches(y + 0.51)); ln.line.color.rgb = LINE; ln.line.width = Pt(0.5); noshadow(ln)
vl = b.shapes.add_connector(1, Inches(6.83), Inches(1.91), Inches(6.83), Inches(4.65)); vl.line.color.rgb = LINE; vl.line.width = Pt(0.75); noshadow(vl)
text(b, 7.09, 1.66, 2.4, 0.22, '더 편한 것을 찾는 고객', 10, True, SUB)
stats = [('4.7배', '숨고 매출', '2021 → 2025'), ('2조 원', 'LG 가전 구독 국내 매출', '2025년, 관리 포함'), ('1,000만 건', '미소 누적 주문', '서비스 1,000개 이상, 2025.9')]
for i, (big, k, d) in enumerate(stats):
    y = 1.95 + i * 0.9
    text(b, 7.09, y, 2.4, 0.5, big, 26, True, INK)
    text(b, 7.09, y + 0.5, 2.4, 0.2, k, 10, True, INK)
    text(b, 7.09, y + 0.68, 2.4, 0.18, d, 8.5, False, SUB)

# 순서: 새 장표 A는 3쪽 뒤, B는 4쪽(원래) 뒤로
ids = prs.slides._sldIdLst; items = list(ids)
newA, newB = items[-2], items[-1]
ids.remove(newA); ids.remove(newB)
ids.insert(3, newA)   # 1,2,3,A
ids.insert(5, newB)   # 1,2,3,A,4,B
prs.save(OUT); print('saved', OUT, len(prs.slides))
