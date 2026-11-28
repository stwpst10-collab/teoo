"""팀 IR 덱(ROOMOVE IR Deck V1) 톤으로 산업 장표 2장을 만들어 덱의 3·4쪽 뒤에 넣는다.
산업 변화가 뒤의 PROBLEM 01~03(판단·조합·실행)과 SOLUTION(홈매니저)으로 이어지도록 구성.
- 덱의 CONTENT 레이아웃(라벨·제목·부제·출처 자리표시자)을 쓰고, 자리표시자 서식은 OVERVIEW 장(7쪽)에서 복사.
사용: python build_ir_market_slides.py <원본 덱.pptx> <출력.pptx>"""
import sys, copy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC, OUT = sys.argv[1], sys.argv[2]
INK = RGBColor(0x11, 0x11, 0x13); SUB = RGBColor(0x6B, 0x6B, 0x70); LIGHT = RGBColor(0xF3, 0xF3, 0xF5)
ACC = RGBColor(0x3D, 0x38, 0xF5); LAV = RGBColor(0xE9, 0xE8, 0xFF); LINE = RGBColor(0xD4, 0xD4, 0xD8)
MUTE = RGBColor(0x9A, 0x9A, 0xA1); WHITE = RGBColor(0xFF, 0xFF, 0xFF); DARK = RGBColor(0x11, 0x11, 0x13)
F = 'Malgun Gothic'

prs = Presentation(SRC)
REF = prs.slides[6]          # OVERVIEW 장(7쪽)
layout = REF.slide_layout     # CONTENT 레이아웃

def text(sl, x, y, w, h, t, size=11, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0; tf.word_wrap = True; tf.vertical_anchor = anchor
    for i, line in enumerate(t.split('\n')):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); para.alignment = align
        r = para.add_run(); r.text = line; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; r.font.name = F
    return tb

def noshadow(shape):
    spPr = shape._element.spPr
    if spPr.find(qn('a:effectLst')) is None: spPr.append(spPr.makeelement(qn('a:effectLst'), {}))

def box(sl, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE, line=None):
    s = sl.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line: s.line.color.rgb = line; s.line.width = Pt(0.75)
    else: s.line.fill.background()
    noshadow(s)
    return s

def hline(sl, x1, y1, x2, y2, color=LINE, w=0.5):
    ln = sl.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2)); ln.line.color.rgb = color; ln.line.width = Pt(w); noshadow(ln)

def fill_ph(sl, label, title, sub, src):
    for ph in sl.placeholders:
        i = ph.placeholder_format.idx
        t = {1: label, 0: title, 2: sub, 3: src}.get(i)
        if t is None: continue
        ph.text_frame.text = t
        refph = [r for r in REF.placeholders if r.placeholder_format.idx == i][0]
        rp = refph.text_frame.paragraphs[0]; para = ph.text_frame.paragraphs[0]
        if rp._p.pPr is not None:
            if para._p.pPr is not None: para._p.remove(para._p.pPr)
            para._p.insert(0, copy.deepcopy(rp._p.pPr))
        if rp.runs and para.runs and rp.runs[0]._r.rPr is not None:
            r = para.runs[0]._r
            if r.rPr is not None: r.remove(r.rPr)
            r.insert(0, copy.deepcopy(rp.runs[0]._r.rPr))

# ── 장표 A: 구조적 변화 → 그런데 고객에게 남은 일(PROBLEM 01~03로 연결) ──
a = prs.slides.add_slide(layout)
fill_ph(a, 'MARKET', '주거 공간 산업 내 구조적인 변화',
        '상품·서비스·공간 제안은 늘었지만, 공간 문제를 판단부터 실행까지 푸는 일은 여전히 고객 몫',
        '출처: 국가데이터처 「온라인쇼핑동향」(가구, 2025 잠정), 한샘·브레이브모바일(숨고) 감사보고서(DART), 유니콘팩토리(2025.9, 미소 회사 발표), 바이라인네트워크(2026.1, 오늘의집 발표), LG전자 실적 발표(2026.1). 오른쪽은 당사 정리')
text(a, 0.56, 1.66, 3.9, 0.22, '시장 내 진행 중인 구조적 변화', 10, True, SUB)
shifts = [('01', '상품 판매의 정체', '상품만 팔아서는 크기 어려운 시장', '온라인 가구 거래액 5조 원대 정체(2021~2025), 한샘 매출 −22%'),
          ('02', '공급의 세분화', '고를 상품과 서비스는 넘치는 시장', '미소 서비스 1,000개 이상, 입점 판매 중심 커머스'),
          ('03', '공간 제안 도구의 확산', '사진·도면으로 미리 보는 집', '오늘의집 3D·AI 인테리어 이용자 연평균 63% 증가'),
          ('04', '맡기는 소비의 성장', '직접 하지 않고 맡기는 집', '숨고 매출 4.7배(2021→2025), LG 가전 구독 2조 원')]
for i, (n, k, v, ev) in enumerate(shifts):
    y = 1.95 + i * 0.73
    box(a, 0.56, y, 3.75, 0.64, LIGHT)
    box(a, 0.7, y + 0.14, 0.36, 0.36, ACC, MSO_SHAPE.OVAL)
    text(a, 0.7, y + 0.14, 0.36, 0.36, n, 9, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(a, 1.18, y + 0.07, 3.05, 0.2, k, 9, True, ACC)
    text(a, 1.18, y + 0.26, 3.05, 0.24, v, 11.5, True, INK)
    text(a, 1.18, y + 0.47, 3.05, 0.16, ev, 7.5, False, MUTE)
box(a, 4.4, 2.95, 0.26, 0.7, ACC, MSO_SHAPE.CHEVRON)
text(a, 4.78, 1.66, 4.7, 0.22, '그런데 고객에게 남은 일', 10, True, SUB)
rows = [('판단', '공간을 보여 주고\n상품을 추천', '무엇을 유지·교체·\n맡길지 판단', 'PROBLEM 01'),
        ('조합', '상품·업체는\n넘치게 많음', '공간에 맞는 조합을\n직접 확인', 'PROBLEM 02'),
        ('실행', '업체마다 따로\n견적·예약', '작업 순서와 일정을\n직접 조율', 'PROBLEM 03')]
text(a, 5.75, 1.9, 1.55, 0.18, '시장이 해 준 것', 8, True, MUTE, PP_ALIGN.CENTER)
text(a, 7.62, 1.9, 1.8, 0.18, '고객에게 남은 일', 8, True, ACC, PP_ALIGN.CENTER)
for i, (k, frm, to, tag) in enumerate(rows):
    y = 2.12 + i * 0.86
    box(a, 4.78, y, 0.85, 0.74, DARK)
    text(a, 4.78, y, 0.85, 0.74, k, 12, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    box(a, 5.75, y, 1.55, 0.74, LIGHT)
    text(a, 5.75, y, 1.55, 0.74, frm, 9.5, True, SUB, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(a, 7.3, y, 0.32, 0.74, '→', 12, True, ACC, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    box(a, 7.62, y, 1.8, 0.74, ACC)
    text(a, 7.62, y + 0.04, 1.8, 0.5, to, 9.5, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(a, 7.62, y + 0.53, 1.8, 0.16, tag, 6.5, True, LAV, PP_ALIGN.CENTER)
text(a, 4.78, 4.72, 4.65, 0.2, '→ 판단부터 실행까지 한 번에 맡을 곳이 필요', 10, True, INK)

# ── 장표 B: 단계별 플레이어 지도(판단·선택·실행) + 더 편한 것을 찾는 고객 ──
b = prs.slides.add_slide(layout)
fill_ph(b, 'MARKET', '기업은 ‘집을 아는 것’, 고객은 ‘더 편한 것’을 찾는 중',
        '판단·선택·실행 단계마다 플레이어는 있지만, 세 단계를 하나로 잇는 곳은 아직 없음',
        '출처: 바이라인네트워크(2026.1)·디지털데일리(2026.3)·벤처스퀘어(2026.10) 오늘의집 보도, 삼성물산 보도자료(2026.11), 브레이브모바일(숨고) 감사보고서, LG전자 실적 발표(2026.1), 유니콘팩토리(2025.9). 분류는 당사 정리(공개 자료 기준)')
steps = [('1  해결 방법 판단', ['오늘의집 AI·3D 방꾸미기', '아키스케치 배치 도구', '쇼룸·라운지 상담(오늘의집·한샘)'], '판매 상품 중심의 제안'),
         ('2  상품·서비스 선택', ['상품: 오늘의집·쿠팡·가전 제조사', '서비스: 숨고·미소·짐싸·헬로입주', '건설사 앱: 홈닉(신축 단지)'], '공간에 맞는 조합은 고객이 확인'),
         ('3  작업 실행·조율', ['오늘의집 직접 시공·지정일 배송', '삼성 인테리어핏(자사 가전)', '이사·청소 업체 개별 예약'], '업종 사이 순서·재조정은 없음')]
x0, w, gap = 0.56, 2.86, 0.14
for i, (h, ps, lim) in enumerate(steps):
    x = x0 + i * (w + gap)
    shp = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
    box(b, x, 1.68, w + (0.12 if i < 2 else 0), 0.42, LIGHT, shp)
    text(b, x + 0.15, 1.68, w - 0.2, 0.42, h, 11, True, INK, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    for j, pl in enumerate(ps):
        yy = 2.24 + j * 0.36
        box(b, x, yy, w, 0.3, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE, LINE)
        text(b, x + 0.12, yy, w - 0.2, 0.3, pl, 9, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    text(b, x, 3.36, w, 0.2, '· ' + lim, 9, False, SUB)
box(b, 0.56, 3.66, 8.86, 0.38, ACC)
text(b, 0.56, 3.66, 8.86, 0.38, '세 단계를 하나의 해결 과정으로 잇는 곳: 공개 자료로 확인된 곳 없음', 10.5, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
text(b, 0.56, 4.14, 2.2, 0.2, '더 편한 것을 찾는 고객', 9, True, SUB)
stats = [('4.7배', '숨고 매출, 2021 → 2025'), ('2조 원', 'LG 가전 구독 국내 매출, 2025'), ('1,000만 건', '미소 누적 주문, 2025.9')]
for i, (big, d) in enumerate(stats):
    x = 2.55 + i * 2.3
    text(b, x, 4.08, 2.2, 0.36, big, 18, True, INK)
    text(b, x, 4.46, 2.2, 0.2, d, 8, False, SUB)

# 순서: A는 3쪽 뒤, B는 원래 4쪽 뒤
ids = prs.slides._sldIdLst; items = list(ids)
newA, newB = items[-2], items[-1]
ids.remove(newA); ids.remove(newB)
ids.insert(3, newA); ids.insert(5, newB)
prs.save(OUT); print('saved', OUT, len(prs.slides))
