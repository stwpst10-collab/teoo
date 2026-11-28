"""IR 덱 톤으로 '통계로 본 타깃 고객의 특징과 페인포인트' 장표 1장을 덱 사본 맨 끝에 추가해 저장한다.
(열린 덱에는 PowerPoint COM InsertFromFile로 넣는다.) 숫자 출처: 리서치/고객 문제 근거 2026-11-25.md, 덱 OVERVIEW 장.
구성은 사용자 예시(K-DUCK 제안 PT: 인물 → 특징 말풍선 / 사선 / 같은 인물 → 페인포인트 말풍선)를 따름.
사용: python build_ir_persona.py <덱.pptx> <출력.pptx>"""
import sys, copy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC, OUT = sys.argv[1], sys.argv[2]
INK = RGBColor(0x11, 0x11, 0x13); SUB = RGBColor(0x6B, 0x6B, 0x70); LIGHT = RGBColor(0xF3, 0xF3, 0xF5)
ACC = RGBColor(0x3D, 0x38, 0xF5); MUTE = RGBColor(0x9A, 0x9A, 0xA1); LINE = RGBColor(0xD4, 0xD4, 0xD8)
LAV = RGBColor(0xE9, 0xE8, 0xFF); WHITE = RGBColor(0xFF, 0xFF, 0xFF); GRAYAV = RGBColor(0xB8, 0xB8, 0xBE)
F = 'Malgun Gothic'
prs = Presentation(SRC)
REF = [s for s in prs.slides if any(ph.placeholder_format.idx == 1 and ph.text_frame.text.strip() == 'OVERVIEW' for ph in s.placeholders)][0]

def noshadow(shape):
    spPr = shape._element.spPr
    if spPr.find(qn('a:effectLst')) is None: spPr.append(spPr.makeelement(qn('a:effectLst'), {}))
def text(sl, x, y, w, h, t, size=11, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, runs=None):
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0; tf.word_wrap = True; tf.vertical_anchor = anchor
    for i, line in enumerate(t.split('\n')):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); para.alignment = align
        r = para.add_run(); r.text = line; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; r.font.name = F
    return tb
def box(sl, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE, line=None):
    s = sl.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h)); s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None: s.line.fill.background()
    else: s.line.color.rgb = line; s.line.width = Pt(1)
    noshadow(s); return s
def conn(sl, x1, y1, x2, y2, color=LINE, w=1):
    c = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color; c.line.width = Pt(w); noshadow(c); return c
def fill_ph(sl, label, title, sub, src):
    for ph in sl.placeholders:
        i = ph.placeholder_format.idx; t = {1: label, 0: title, 2: sub, 3: src}.get(i)
        if t is None: continue
        ph.text_frame.text = t
        rp = [r for r in REF.placeholders if r.placeholder_format.idx == i][0].text_frame.paragraphs[0]; para = ph.text_frame.paragraphs[0]
        if rp._p.pPr is not None:
            if para._p.pPr is not None: para._p.remove(para._p.pPr)
            para._p.insert(0, copy.deepcopy(rp._p.pPr))
        if rp.runs and para.runs and rp.runs[0]._r.rPr is not None:
            r = para.runs[0]._r
            if r.rPr is not None: r.remove(r.rPr)
            r.insert(0, copy.deepcopy(rp.runs[0]._r.rPr))
def avatar(sl, x, y, d, fill, name, desc):
    box(sl, x, y, d, d, fill, MSO_SHAPE.OVAL)
    text(sl, x, y + d * 0.30, d, 0.3, name, 15, True, WHITE, PP_ALIGN.CENTER)
    text(sl, x, y + d * 0.30 + 0.34, d, 0.22, desc, 7.5, False, WHITE, PP_ALIGN.CENTER)

s = prs.slides.add_slide(REF.slide_layout)
fill_ph(s, 'PERSONA', '통계로 본 타깃 고객의 특징과 페인포인트',
        '자주 옮기고 오래된 아파트에 사는 가구, 준비의 판단·선택·실행 모든 단계에서 부담',
        '출처: 국토교통부 「2025년도 주거실태조사」, 국가데이터처 「2025년 인구주택총조사」, 엠브레인 트렌드모니터(2025.10·2019.12, 각 1,000명), 한국소비자원 자료 인용 보도(SBS Biz 2026.10, 경향신문 2026.6). 문제 단계 연결은 당사 정리')
text(s, 0.56, 1.62, 4.3, 0.3, '누가 이사·공간 변화를 겪나', 11, True, ACC)
AX, AY, AD = 0.62, 2.55, 1.45
avatar(s, AX, AY, AD, ACC, '타깃 가구', '이사·입주를 앞둔 가구')
traits = [('31.2%', '지금 집에 산 지\n2년 이내인 가구'), ('65.8%', '전체 주택 중\n아파트 비중'),
          ('30.6%', '지은 지 30년이\n넘은 주택'), ('39.9%', '최근 1년 홈인테리어\n경험(셀프 66.1%)')]
bx, bw, bh = 2.55, 2.2, 0.58
for i, (n, t) in enumerate(traits):
    by = 2.05 + i * 0.7
    conn(s, AX + AD - 0.05, AY + AD / 2, bx, by + bh / 2)
    b = box(s, bx, by, bw, bh, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE, ACC); b.adjustments[0] = 0.5
    text(s, bx + 0.14, by, 0.72, bh, n, 12, True, ACC, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(s, bx + 0.9, by, bw - 1.0, bh, t, 7.5, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
conn(s, 5.2, 4.85, 4.95, 1.75, ACC, 1.5)
RX, RY, RD = 5.45, 2.0, 1.45
avatar(s, RX, RY, RD, GRAYAV, '준비 과정', '판단 → 선택 → 실행')
pains = [('PROBLEM 01 · 판단', '80.9%', '이사 준비 과정이\n복잡하고 어렵다'),
         ('PROBLEM 02 · 선택·조합', '77.8%', '이사업체를 고를 때\n여러 곳을 직접 비교'),
         ('PROBLEM 03 · 실행', '25.2%', '인테리어 피해구제 중\n계약 불이행\n(공사 기간 미준수 등)')]
px, pw, ph = 7.08, 2.35, 0.78
for i, (k, n, t) in enumerate(pains):
    py = 1.72 + i * 0.88
    conn(s, RX + RD - 0.05, RY + RD / 2, px, py + ph / 2)
    box(s, px, py, pw, ph, LIGHT)
    text(s, px + 0.12, py + 0.08, pw - 0.24, 0.18, k, 7, True, ACC)
    text(s, px + 0.12, py + 0.28, 0.72, 0.44, n, 13, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    text(s, px + 0.84, py + 0.28, pw - 0.92, 0.44, t, 7, False, SUB, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
text(s, 5.45, 4.42, 3.98, 0.4, '온라인 가구 피해구제의 26.4%도 배송 지연·미배송(2025)', 9.5, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
prs.save(OUT); print('saved', OUT, len(prs.slides))
