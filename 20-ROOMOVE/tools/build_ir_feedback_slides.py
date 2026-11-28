"""2026-11-28 팀 통화 피드백 반영 장표 2장을 덱 사본 맨 끝에 추가한다(IR 덱 톤).
  1) MARKET: 전통 강자는 줄고 플랫폼 강자는 커지는 흐름(2021=100 지수 꺾은선)
  2) COMPETITION: 단계별 대표 플레이어만 골라 규모로 보기(오늘의집은 판단·상품 선택을 하나로 묶음)
숫자: 리서치/확장/주거 산업 시계열 2026-11-27.md (DART 감사보고서·공시), 바이라인네트워크 2026.1(오늘의집 발표)
사용: python build_ir_feedback_slides.py <덱.pptx> <출력.pptx>"""
import sys, copy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.oxml.ns import qn

SRC, OUT = sys.argv[1], sys.argv[2]
INK = RGBColor(0x11, 0x11, 0x13); SUB = RGBColor(0x6B, 0x6B, 0x70); LIGHT = RGBColor(0xF3, 0xF3, 0xF5)
ACC = RGBColor(0x3D, 0x38, 0xF5); MUTE = RGBColor(0x9A, 0x9A, 0xA1); LINE = RGBColor(0xD4, 0xD4, 0xD8)
LAV = RGBColor(0xE9, 0xE8, 0xFF); WHITE = RGBColor(0xFF, 0xFF, 0xFF); DARK = RGBColor(0x2A, 0x2A, 0x2E)
ACC2 = RGBColor(0x8E, 0x8A, 0xFF); GRAY1 = RGBColor(0x9A, 0x9A, 0xA1); GRAY2 = RGBColor(0xC4, 0xC4, 0xCA); RED = RGBColor(0xC1, 0x3A, 0x13)
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
def hline(sl, x1, y, x2, color=LINE):
    ln = sl.shapes.add_connector(1, Inches(x1), Inches(y), Inches(x2), Inches(y)); ln.line.color.rgb = color; ln.line.width = Pt(0.75); noshadow(ln)
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

# ---------- 1) 전통 강자 vs 플랫폼 강자 ----------
s = prs.slides.add_slide(REF.slide_layout)
fill_ph(s, 'MARKET', '전통 강자는 줄고, 플랫폼 강자는 커지는 주거 서비스',
        '같은 4년 동안 가구 대기업·부동산 중개 앱 매출은 줄고, 공간 제안·홈서비스 플랫폼 매출은 3~5배로 성장',
        '출처: 각 사 감사보고서·공시(DART). 한샘 연결, 다방(스테이션3)·오늘의집(버킷플레이스)·숨고(브레이브모바일) 별도 매출. 지수는 2021년=100 당사 계산(한샘 2025 잠정)')
series = [('숨고', [151.8, 313.8, 458.6, 624.2, 719.8], ACC, True), ('오늘의집', [1176.3, 1864.2, 2355.2, 2879.5, 3215.7], ACC2, True),
          ('다방', [246.0, 230.7, 209.4, 209.7, 206.5], GRAY2, False), ('한샘', [22314, 20009, 19669, 19084, 17445], GRAY1, False)]
cd = CategoryChartData(); cd.categories = ['2021', '2022', '2023', '2024', '2025']
for nm, v, c, up in series: cd.add_series(nm, [round(x / v[0] * 100, 1) for x in v])
gf = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(0.5), Inches(1.6), Inches(5.6), Inches(3.3), cd)
ch = gf.chart; ch.has_legend = False; ch.font.name = F; ch.font.size = Pt(8); ch.font.color.rgb = SUB
va = ch.value_axis; va.has_major_gridlines = True; va.major_gridlines.format.line.color.rgb = RGBColor(0xEC, 0xEC, 0xEE)
va.format.line.fill.background(); va.tick_labels.font.size = Pt(7.5); va.maximum_scale = 500; va.minimum_scale = 0; va.major_unit = 100
ca = ch.category_axis; ca.format.line.color.rgb = LINE; ca.tick_labels.font.size = Pt(8)
for i, (nm, v, c, up) in enumerate(series):
    ser = ch.plots[0].series[i]; ser.smooth = False
    ser.format.line.color.rgb = c; ser.format.line.width = Pt(2.5 if up else 1.75)
    ser.marker.style = XL_MARKER_STYLE.CIRCLE; ser.marker.size = 5
    ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = c; ser.marker.format.line.color.rgb = c
    pt = ser.points[4]; pt.data_label.has_text_frame = True
    pt.data_label.text_frame.text = f'{nm} {round(v[4] / v[0] * 100)}'
    r = pt.data_label.text_frame.paragraphs[0].runs[0]; r.font.size = Pt(8); r.font.bold = True; r.font.color.rgb = c if up else SUB; r.font.name = F
    from pptx.enum.chart import XL_LABEL_POSITION
    pt.data_label.position = XL_LABEL_POSITION.RIGHT if up else (XL_LABEL_POSITION.ABOVE if nm == '다방' else XL_LABEL_POSITION.BELOW)
text(s, 0.56, 1.58, 3, 0.2, '매출 지수(2021년=100)', 8, True, SUB)
# 오른쪽 카드
cx, cw = 6.35, 3.08
text(s, cx, 1.62, cw, 0.22, '2025년 매출과 2021년 대비 변화', 9, True, INK)
cards = [('플랫폼', '숨고', '홈서비스 연결', '720억 원', '4.7배', True), ('플랫폼', '오늘의집', '공간 제안·커머스·시공', '3,216억 원', '2.7배', True),
         ('전통', '다방', '부동산 중개 앱', '207억 원', '−16%', False), ('전통', '한샘', '가구·인테리어 대기업', '1조 7,445억 원', '−22%', False)]
for j, (kind, nm, desc, rev, chg, up) in enumerate(cards):
    y = 1.92 + j * 0.74
    box(s, cx, y, cw, 0.66, LAV if up else LIGHT)
    text(s, cx + 0.14, y + 0.1, 0.5, 0.18, kind, 7, True, ACC if up else MUTE)
    text(s, cx + 0.14, y + 0.27, 1.3, 0.26, nm, 11, True, INK)
    text(s, cx + 0.14, y + 0.48, 1.5, 0.16, desc, 6.5, False, SUB)
    text(s, cx + 1.35, y + 0.12, 1.0, 0.44, rev, 8.5, False, INK, PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
    text(s, cx + 2.4, y + 0.12, 0.56, 0.44, chg, 13, True, ACC if up else RED, PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)

# ---------- 2) 단계별 대표 플레이어 ----------
s = prs.slides.add_slide(REF.slide_layout)
fill_ph(s, 'COMPETITION', '단계마다 강자는 있지만, 세 단계를 잇는 곳은 없음',
        '판단·상품 선택은 오늘의집, 서비스 선택은 숨고, 시공 실행은 한샘이 대표. 각자 자기 단계에서만 강함',
        '출처: 각 사 감사보고서·공시(DART, 2025년 매출. 한샘 연결·잠정, 오늘의집·숨고 별도), 바이라인네트워크(2026.1, 오늘의집 발표), 각 사 서비스 소개. 단계 분류와 한계는 당사 정리(공개 자료 기준)')
x0, y0, H = 0.56, 1.62, 2.72
chev = [('1  해결 방법 판단', 0.56, 2.9), ('2  상품·서비스 선택', 3.38, 2.9), ('3  작업 실행·조율', 6.53, 2.9)]
# 단계 띠
for k, (lab, x, w) in enumerate([('1  해결 방법 판단', 0.56, 2.5), ('2  상품·서비스 선택', 3.0, 4.38), ('3  작업 실행·조율', 7.32, 2.11)]):
    sh = s.shapes.add_shape(MSO_SHAPE.CHEVRON if k else MSO_SHAPE.PENTAGON, Inches(x), Inches(y0), Inches(w), Inches(0.32))
    sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor(0xE4, 0xE4, 0xE7); sh.line.fill.background(); noshadow(sh)
    text(s, x + 0.2, y0, w - 0.4, 0.32, lab, 9, True, INK, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
players = [
    (0.56, 4.6, '오늘의집', '판단 + 상품 선택을 한 앱에서', '3,216억 원', '2025 매출 · 2021년 대비 2.7배',
     ['3D·AI 방꾸미기 사용자 연평균 63% 증가(2023~)', '공간 제안 → 상품 판매 → 직접 시공까지'],
     '판매 상품 중심의 제안, 이사·청소 등 다른 업종 일정은 고객 몫', True),
    (5.3, 1.92, '숨고', '서비스 선택', '720억 원', '2025 매출 · 4.7배',
     ['이사·청소·시공 업체 견적 연결', '2023년부터 흑자'],
     '업체별 견적까지, 공간 판단과 작업 순서는 없음', False),
    (7.36, 2.07, '한샘', '작업 실행(시공)', '1조 7,445억 원', '2025 매출 · −22%',
     ['리하우스 시공·자사 가구 설치'],
     '자사 상품 중심, 다른 업체 일정은 고객 몫', False),
]
py = y0 + 0.46
for (x, w, nm, role, rev, revsub, pts, lim, big) in players:
    box(s, x, py, w, 2.36, LAV if big else LIGHT)
    text(s, x + 0.16, py + 0.14, 2.2, 0.18, role, 7.5, True, ACC)
    text(s, x + 0.16, py + 0.34, 2.2, 0.3, nm, 14, True, INK)
    text(s, x + 0.16, py + 0.7, 2.4, 0.34, rev, 17 if big else 14, True, INK)
    text(s, x + 0.16, py + 1.04, 2.2, 0.18, revsub, 7, False, SUB)
    if big:
        tx, ty, tw = x + 2.35, py + 0.34, w - 2.5
    else:
        hline(s, x + 0.16, py + 1.3, x + w - 0.16); tx, ty, tw = x + 0.16, py + 1.38, w - 0.32
    text(s, tx, ty, tw, 0.5, '\n'.join('· ' + p for p in pts), 7.5, False, INK)
    text(s, tx, ty + (0.72 if big else 0.48), tw, 0.46, '한계  ' + lim, 7.5, True, RED)
by = py + 2.36 + 0.12
box(s, 0.56, by, 8.87, 0.4, ACC)
text(s, 0.76, by, 8.5, 0.4, '세 단계를 한 해결 과정으로 잇는 곳: 대중형 서비스로는 공개 자료상 없음(사람이 하는 고급 대행 무브제 정도)', 9.5, True, WHITE, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
prs.save(OUT); print('saved', OUT, len(prs.slides))
