"""IR 덱 톤으로 '타깃 페르소나 특징과 페인포인트' 장표 1장을 덱 사본 맨 끝에 추가해 저장한다.
(열린 덱에는 PowerPoint COM InsertFromFile로 넣는다.) 인물은 덱 TARGET·PROBLEM 장 기준 가상 인물.
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
fill_ph(s, 'PERSONA', '타깃 페르소나 특징과 페인포인트',
        '이사 두 달 전, 판단부터 일정까지 혼자 떠안은 준비 담당자',
        '가상 인물. 덱 TARGET 장의 초기 타깃(이사 1~3개월 전, 쓰던 가구와 신규 구매, 작업 2건 이상)과 PROBLEM 장의 제품 시나리오를 한 가구로 그린 것으로, 실제 인터뷰 결과 아님')

# 왼쪽: 특징
text(s, 0.56, 1.62, 4.2, 0.3, '“새집을 우리 가족에게 맞게 바꾸고 싶은데”', 12, True, ACC)
AX, AY, AD = 0.62, 2.55, 1.45
avatar(s, AX, AY, AD, ACC, '지현', '38세 · 맞벌이 부모')
traits = ['구축 아파트 매수,\n입주까지 8주', '초등 1학년 자녀,\n아이 방을 새로 꾸밀 계획',
          '소파·TV장·침대는\n새집에서도 그대로', '시공·이사·청소·설치\n맡길 작업 4건']
bx, bw, bh = 2.55, 2.1, 0.58
for i, t in enumerate(traits):
    by = 2.05 + i * 0.7
    conn(s, AX + AD - 0.05, AY + AD / 2, bx, by + bh / 2)
    b = box(s, bx, by, bw, bh, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE, ACC); b.adjustments[0] = 0.5
    text(s, bx, by, bw, bh, t, 8.5, True, INK, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)

# 가운데 사선
conn(s, 5.2, 4.85, 4.95, 1.75, ACC, 1.5)

# 오른쪽: 페인포인트
RX, RY, RD = 5.45, 2.0, 1.45
avatar(s, RX, RY, RD, GRAYAV, '지현', '이사 준비 6주째')
pains = [('PROBLEM 01 · 판단', '침대를 두고 책상만 살지,\n붙박이장을 뗄지 정하기 어려움'),
         ('PROBLEM 02 · 조합', '새 책상·수납장이 쓰던 침대 옆에\n들어가는지 직접 확인'),
         ('PROBLEM 03 · 실행', '철거가 하루 밀리면 청소·배송 일정을\n업체마다 다시 연락')]
px, pw, ph = 7.08, 2.35, 0.78
for i, (k, t) in enumerate(pains):
    py = 1.72 + i * 0.88
    conn(s, RX + RD - 0.05, RY + RD / 2, px, py + ph / 2)
    box(s, px, py, pw, ph, LIGHT)
    text(s, px + 0.12, py + 0.09, pw - 0.24, 0.18, k, 7, True, ACC)
    text(s, px + 0.12, py + 0.3, pw - 0.2, 0.42, t, 8, True, INK)
text(s, 5.45, 4.42, 3.98, 0.4, '“고를 건 넘치는데, 우리 집에 맞게 정해 줄 곳이 없네”', 11.5, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
prs.save(OUT); print('saved', OUT, len(prs.slides))
