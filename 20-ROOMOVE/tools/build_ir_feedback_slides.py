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

# ---------- 2) 단계별 경쟁사 지도 + 대표 경쟁사 강조(매출 막대) ----------
s = prs.slides.add_slide(REF.slide_layout)
fill_ph(s, 'COMPETITION', '단계마다 강자는 있지만, 세 단계를 잇는 곳은 없음',
        '판단·상품 선택은 오늘의집, 서비스 선택은 숨고, 시공 실행은 한샘이 대표. 오늘의집은 세 단계로 넓히는 중',
        '출처: 각 사 감사보고서·공시(DART, 2025년 매출. 한샘 연결·잠정, 오늘의집·숨고·영구크린 별도), 미소는 회사 발표 최근 1년 거래액(유니콘팩토리 2025.7, 매출 아님), 바이라인네트워크(2026.1)·디지털데일리(2026.3)·삼성물산 보도자료(2026.11). 막대는 회사별 자기 매출 추이(회사 간 비교 아님). 분류는 당사 정리')
TREND = {'오늘의집': [1176, 1864, 2355, 2880, 3216], '숨고': [152, 314, 459, 624, 720], '한샘': [22314, 20009, 19669, 19084, 17445]}
cols = [
    ('1  해결 방법 판단', ('오늘의집', '3D·AI 방꾸미기 · 쇼룸 상담', 3216, '3,216억', '2.7배'),
     [('아키스케치', '배치 도구', '매출 비공개', False), ('한샘 쇼룸', '매장 상담 → 3단계 한샘', '', False)]),
    ('2  상품·서비스 선택', ('숨고', '이사·청소·시공 업체 견적 연결', 720, '720억', '4.7배 · 2023년부터 흑자'),
     [('오늘의집', '상품 커머스(1단계와 같은 회사)', '', True), ('미소', '생활서비스', '연 거래 약 5,000억', False),
      ('쿠팡·가전 제조사', '상품 판매', '', False), ('짐싸·헬로입주·홈닉', '이사·입주 서비스', '매출 비공개', False)]),
    ('3  작업 실행·조율', ('한샘', '리하우스 시공 · 자사 가구 설치', 17445, '1조 7,445억', '−22%'),
     [('오늘의집', '직접 시공 · 지정일 배송', '', True), ('영구크린', '이사(가맹)', '매출 205억', False),
      ('삼성 인테리어핏', '자사 가전 중심 인테리어', '', False), ('이사·청소 업체', '개별 예약', '', False)]),
]
cw, gap, x0, y0 = 2.85, 0.16, 0.56, 1.6
for k, (stage, rep, others) in enumerate(cols):
    x = x0 + k * (cw + gap)
    sh = s.shapes.add_shape(MSO_SHAPE.CHEVRON if k else MSO_SHAPE.PENTAGON, Inches(x), Inches(y0), Inches(cw + (0.12 if k < 2 else 0)), Inches(0.3))
    sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor(0xE4, 0xE4, 0xE7); sh.line.fill.background(); noshadow(sh)
    text(s, x + 0.2, y0, cw - 0.3, 0.3, stage, 9, True, INK, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    # 대표 경쟁사 카드(강조)
    nm, desc, rev, revtxt, chg = rep; ry = y0 + 0.4
    box(s, x, ry, cw, 1.0, ACC)
    text(s, x + 0.14, ry + 0.07, 1.0, 0.18, '대표 경쟁사', 7, True, RGBColor(0xC9, 0xC7, 0xFF))
    text(s, x + 0.14, ry + 0.22, 1.5, 0.3, nm, 14, True, WHITE)
    text(s, x + 0.14, ry + 0.5, cw - 0.28, 0.18, desc, 7, False, RGBColor(0xE0, 0xDF, 0xFF))
    text(s, x + 1.45, ry + 0.16, cw - 1.59, 0.34, revtxt, 14, True, WHITE, PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
    # 회사별 자기 매출 추이(2021~2025, 각자 기준. 회사끼리 비교하지 않음)
    ser = TREND[nm]; vmax = max(ser); sw = (cw - 0.28) / 5; bh = 0.2
    for i, v in enumerate(ser):
        hh = bh * v / vmax
        box(s, x + 0.14 + i * sw + sw * 0.15, ry + 0.66 + bh - hh, sw * 0.7, hh, WHITE if i == 4 else RGBColor(0x8E, 0x8A, 0xFF))
    text(s, x + 0.14, ry + 0.87, cw - 0.28, 0.12, f"'21→'25 매출 추이 · {chg}", 6.5, False, RGBColor(0xE0, 0xDF, 0xFF))
    # 다른 경쟁사(회색)
    oy = ry + 1.06
    text(s, x, oy, cw, 0.18, '그 밖의 플레이어', 7, True, MUTE)
    oy += 0.2
    for (onm, odesc, orev, same) in others:
        box(s, x, oy, cw, 0.27, LAV if same else LIGHT)
        text(s, x + 0.12, oy, 1.15, 0.27, onm, 7.5, True, ACC if same else INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
        text(s, x + 1.25, oy, cw - 1.35, 0.27, odesc + (f' · {orev}' if orev else ''), 6.5, False, SUB, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
        oy += 0.3
by = 4.5
box(s, 0.56, by, 8.87, 0.36, DARK)
text(s, 0.76, by, 8.5, 0.36, '세 단계를 한 해결 과정으로 잇는 곳: 대중형 서비스로는 공개 자료상 없음(사람이 하는 고급 대행 무브제 정도)', 9, True, WHITE, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
prs.save(OUT); print('saved', OUT, len(prs.slides))
