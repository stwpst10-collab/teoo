"""IR 덱 톤으로 '숫자로 본 시장 내 구조적 변화' 차트 장표 1장을 덱 사본 맨 끝에 추가해 저장한다.
(열린 덱에는 PowerPoint COM InsertFromFile로 넣는다.) 숫자 출처: 리서치/확장/주거 산업 시계열 2026-11-27.md
사용: python build_ir_market_charts.py <덱.pptx> <출력.pptx>"""
import sys, copy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.oxml.ns import qn

SRC, OUT = sys.argv[1], sys.argv[2]
INK = RGBColor(0x11, 0x11, 0x13); SUB = RGBColor(0x6B, 0x6B, 0x70); LIGHT = RGBColor(0xF3, 0xF3, 0xF5)
ACC = RGBColor(0x3D, 0x38, 0xF5); MUTE = RGBColor(0x9A, 0x9A, 0xA1); GRAYBAR = RGBColor(0xB8, 0xB8, 0xBE); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
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
def box(sl, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE):
    s = sl.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h)); s.fill.solid(); s.fill.fore_color.rgb = fill; s.line.fill.background(); noshadow(s); return s
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
fill_ph(s, 'MARKET', '숫자로 본 시장 내 구조적 변화',
        '상품 판매는 멈추고, 서비스 공급·공간 제안·맡기는 소비는 빠르게 커지는 흐름',
        '출처: 국가데이터처 「온라인쇼핑동향」(가구, 2025 잠정), 영구크린 감사보고서(DART), 디지털데일리(2026.1, 청소연구소 발표), 유니콘팩토리·회사 연혁(미소 누적 거래액, 회사 발표), 바이라인네트워크(2026.1, 오늘의집 발표: 3D·AI 인테리어 관련 검색량 전년 대비 약 24배), 브레이브모바일(숨고) 감사보고서(DART)')
panels = [
    ('01', '상품 판매의 정체', '2021년부터 정체', '온라인 가구 거래액(조 원)', ['18', '19', '20', '21', '22', '23', '24', '25'], [3.13, 3.54, 4.90, 5.40, 5.20, 5.35, 5.67, 5.53], False, '0.0'),
    ('02', '공급의 세분화', '서비스별 1.8~4.3배', '2022년 대비 배수(2022=1)', None, None, True, None),
    ('03', '공간 제안 도구의 확산', '검색량 1년 새 약 24배', '3D·AI 인테리어 검색량(전년=1)', ['2024', '2025'], [1, 24], True, '0'),
    ('04', '맡기는 소비의 성장', '4년간 4.7배', '숨고 매출(억 원)', ['21', '22', '23', '24', '25'], [151.8, 313.8, 458.6, 624.2, 719.8], True, '#,##0'),
]
x0, gap = 0.56, 0.16; w = (8.87 - gap * 3) / 4
for i, (n, k, take, cap, cats, vals, up, fmt) in enumerate(panels):
    x = x0 + i * (w + gap); y = 1.68
    box(s, x, y, w, 3.18, LIGHT)
    c = box(s, x + 0.12, y + 0.12, 0.3, 0.3, ACC, MSO_SHAPE.OVAL)
    text(s, x + 0.12, y + 0.12, 0.3, 0.3, n, 8, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(s, x + 0.5, y + 0.14, w - 0.6, 0.26, k, 9, True, ACC, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    text(s, x + 0.12, y + 0.5, w - 0.24, 0.3, take, 11.5, True, INK)
    if cats is None:
        rows = [('이사', "영구크린 매출 '22→'25", 1.8), ('청소', "청소연구소 매니저 수 '22→'26.1", 2.2), ('생활서비스', "미소 누적 거래액 '22.4→'25 중반", 4.3)]
        maxw = w - 0.24 - 0.55
        for j, (nm, met, v) in enumerate(rows):
            ry = y + 1.0 + j * 0.6
            text(s, x + 0.12, ry, w - 0.24, 0.2, nm, 9, True, INK)
            text(s, x + 0.12, ry + 0.17, w - 0.24, 0.16, met, 6.5, False, SUB)
            bl = maxw * v / 4.3
            box(s, x + 0.12, ry + 0.36, bl, 0.18, ACC)
            text(s, x + 0.12 + bl + 0.06, ry + 0.32, 0.5, 0.26, f'{v:.1f}배', 9, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
        text(s, x + 0.12, y + 2.85, w - 0.24, 0.26, cap, 7.5, False, SUB)
        continue
    cd = CategoryChartData(); cd.categories = cats; cd.add_series(cap, vals)
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x + 0.05), Inches(y + 0.85), Inches(w - 0.1), Inches(1.95), cd)
    ch = gf.chart; ch.has_legend = False; ch.has_title = False
    ch.font.name = F; ch.font.size = Pt(7.5); ch.font.color.rgb = SUB
    va = ch.value_axis; va.visible = False; va.has_major_gridlines = False
    ca = ch.category_axis; ca.format.line.color.rgb = RGBColor(0xD4, 0xD4, 0xD8); ca.tick_labels.font.size = Pt(7.5)
    pl = ch.plots[0]; pl.gap_width = 45; pl.has_data_labels = True
    dl = pl.data_labels; dl.number_format = fmt; dl.number_format_is_linked = False; dl.position = XL_LABEL_POSITION.OUTSIDE_END; dl.font.size = Pt(7.5); dl.font.color.rgb = INK
    ser = pl.series[0]; ser.format.fill.solid(); ser.format.fill.fore_color.rgb = ACC if up else GRAYBAR
    text(s, x + 0.12, y + 2.85, w - 0.24, 0.26, cap, 7.5, False, SUB)
prs.save(OUT); print('saved', OUT, len(prs.slides))
