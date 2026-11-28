"""IR 덱 톤으로 '가구 유형별 페르소나' 장표 1장을 덱 사본 맨 끝에 추가해 저장한다.
2026-11-28 팀 통화 피드백: 페르소나는 고객 유형(가족·1인·고령)별 고민을 보여 주고, 뒤 OVERVIEW(공간을 바꿔야 하는 순간)와 이어질 것.
숫자: 국가데이터처 2025 인구주택총조사(2026.7.28 발표), 소비자원 소규모 이사 피해구제(헤럴드경제 2026.4), 덱 OVERVIEW 장.
사용: python build_ir_persona_types.py <덱.pptx> <출력.pptx>"""
import sys, copy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC, OUT = sys.argv[1], sys.argv[2]
INK = RGBColor(0x11, 0x11, 0x13); SUB = RGBColor(0x6B, 0x6B, 0x70); LIGHT = RGBColor(0xF3, 0xF3, 0xF5)
ACC = RGBColor(0x3D, 0x38, 0xF5); MUTE = RGBColor(0x9A, 0x9A, 0xA1); LINE = RGBColor(0xD4, 0xD4, 0xD8)
LAV = RGBColor(0xE9, 0xE8, 0xFF); WHITE = RGBColor(0xFF, 0xFF, 0xFF); DARK = RGBColor(0x2A, 0x2A, 0x2E)
F = 'Malgun Gothic'
prs = Presentation(SRC)
REF = [s for s in prs.slides if any(ph.placeholder_format.idx == 1 and ph.text_frame.text.strip() == 'OVERVIEW' for ph in s.placeholders)][0]

def noshadow(shape):
    spPr = shape._element.spPr
    if spPr.find(qn('a:effectLst')) is None: spPr.append(spPr.makeelement(qn('a:effectLst'), {}))
def text(sl, x, y, w, h, t, size=11, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0; tf.word_wrap = True; tf.vertical_anchor = anchor
    for i, line in enumerate(t.split('\n')):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); para.alignment = align
        r = para.add_run(); r.text = line; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; r.font.name = F
    return tb
def box(sl, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE):
    s = sl.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h)); s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background(); noshadow(s); return s
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

s = prs.slides.add_slide(REF.slide_layout)
fill_ph(s, 'PERSONA', '가구 유형마다 다른 계기, 같은 공간 고민',
        '공간을 바꾸는 계기는 유형마다 다르지만, 판단·조합·실행을 직접 떠안는 부담은 같음',
        '출처: 국가데이터처 「2025년 인구주택총조사」(2026.7.28 발표, 일반가구 2,249.5만 기준, 4인 이상 비중은 당사 계산), 한국소비자원 자료 인용 보도(헤럴드경제 2026.4, 2025년 소규모 이사 피해구제 241건). 인물 예시는 제품 시나리오(가상)')
types = [
    ('01', '아이 있는 가족', '예) 38세 맞벌이 부부, 구축 아파트로 이사', '340.8만', '4인 이상 가구 (일반가구의 15.2%)',
     '“쓰던 침대는 두고 책상만 살까,\n붙박이장을 뗄까?”', ['이사·입주', '자녀 성장'], '시공·이사·청소·설치 4건의 순서와 일정'),
    ('02', '1인 가구', '예) 29세 직장인, 2년마다 원룸·오피스텔 이사', '824.4만', '1인 가구 (일반가구의 36.6%, 역대 최대)',
     '“작은 집에 짐이 다 들어갈까,\n버릴까 맡길까?”', ['이사·입주', '수납 부족'], '혼자 업체 비교, 당일 추가 비용 대응\n(소규모 이사 피해 65.1%가 20~30대)'),
    ('03', '고령 가구', '예) 60대 후반 부부, 30년 된 아파트에 거주', '245.6만', '65세 이상 1인 가구 (일반가구의 10.9%, 역대 최대)',
     '“욕실은 전체 공사가 필요할까,\n부분 보수로 충분할까?”', ['욕실 노후화', '작은 집으로 이사'], '공사 범위 판단과 공사 중 생활 불편'),
]
x0, gap, y0, H = 0.56, 0.16, 1.62, 2.72
w = (8.87 - gap * 2) / 3
for i, (n, name, ex, num, numlab, quote, moments, burden) in enumerate(types):
    x = x0 + i * (w + gap)
    box(s, x, y0, w, H, LIGHT)
    box(s, x + 0.16, y0 + 0.16, 0.3, 0.3, ACC, MSO_SHAPE.OVAL)
    text(s, x + 0.16, y0 + 0.16, 0.3, 0.3, n, 8, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(s, x + 0.56, y0 + 0.15, w - 0.7, 0.32, name, 13, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    text(s, x + 0.16, y0 + 0.52, w - 0.32, 0.2, ex, 7.5, False, SUB)
    text(s, x + 0.16, y0 + 0.76, w - 0.32, 0.4, num, 22, True, INK)
    text(s, x + 0.16, y0 + 1.16, w - 0.32, 0.2, numlab, 7.5, False, SUB)
    ln = s.shapes.add_connector(1, Inches(x + 0.16), Inches(y0 + 1.43), Inches(x + w - 0.16), Inches(y0 + 1.43)); ln.line.color.rgb = LINE; ln.line.width = Pt(0.75); noshadow(ln)
    text(s, x + 0.16, y0 + 1.52, w - 0.32, 0.42, quote, 9.5, True, INK)
    text(s, x + 0.16, y0 + 2.0, 0.6, 0.22, '계기', 7.5, True, ACC, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    px = x + 0.62
    for m in moments:
        pw = 0.16 + len(m) * 0.105
        p = box(s, px, y0 + 2.0, pw, 0.22, LAV, MSO_SHAPE.ROUNDED_RECTANGLE); p.adjustments[0] = 0.5
        text(s, px, y0 + 2.0, pw, 0.22, m, 7.5, True, ACC, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        px += pw + 0.08
    text(s, x + 0.16, y0 + 2.3, 0.6, 0.2, '부담', 7.5, True, ACC)
    text(s, x + 0.62, y0 + 2.3, w - 0.78, 0.36, burden, 7.5, False, INK)
by = y0 + H + 0.14
box(s, x0, by, 8.87, 0.42, DARK)
text(s, x0 + 0.2, by, 4.6, 0.42, '유형은 달라도 고객에게 남는 일은 같음', 10, True, WHITE, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
for j, (k, tag) in enumerate([('판단', 'PROBLEM 01'), ('조합', 'PROBLEM 02'), ('실행', 'PROBLEM 03')]):
    cx = x0 + 4.55 + j * 1.43
    text(s, cx, by, 1.3, 0.42, f'{k}  ', 10, True, WHITE, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    text(s, cx + 0.42, by, 0.9, 0.42, tag, 7, True, RGBColor(0xB9, 0xB6, 0xFF), PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
prs.save(OUT); print('saved', OUT, len(prs.slides))
