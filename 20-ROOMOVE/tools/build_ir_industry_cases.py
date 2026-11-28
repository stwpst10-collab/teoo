"""2026-11-28 IR 덱 MARKET 장 1장: 업계 확장 사례(덱 사본 맨 끝에 추가).
앞 장(build_ir_industry_keywords.py)의 '집을 이해하고 맡아 주는 쪽' 흐름을 회사 사례 6개로 받친다.
회사별 한 줄: 회사·본업(회색) → 넓힌 영역 칩(시점순, 가장 최근은 파랑) → 확인된 숫자 1개.
숫자·출처: 리서치/확장/업계 확장 사례 2026-11-28.md
톤: build_ir_industry_keywords.py와 같은 헬퍼·색 값, OVERVIEW 장의 CONTENT 레이아웃·자리표시자 서식 복사
사용: python build_ir_industry_cases.py <덱.pptx> <출력.pptx>"""
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
PALE = RGBColor(0xC9, 0xC7, 0xFF)
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
def rich(sl, x, y, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE):
    """lines: [(글, 크기, 굵게, 색)] 한 문단씩"""
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0; tf.word_wrap = True; tf.vertical_anchor = anchor
    for i, (t, size, bold, color) in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); para.alignment = align
        r = para.add_run(); r.text = t; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; r.font.name = F
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
fill_ph(s, 'MARKET', "본업에서 출발해 '집 전체'를 맡는 쪽으로 넓히는 업계",
        '인테리어·가구·건설·가전·부동산·이사 업계 모두 판매 다음 단계인 시공·관리·입주 서비스로 영역을 넓히는 흐름',
        '출처: 버킷플레이스·직방 감사보고서(DART), 삼성물산·LG전자·무브제 발표, 바이라인네트워크(2026.10)·지디넷(2026.7)·CEO스코어데일리(2025.9·2026.11)·쿠키뉴스(2026.6)·'
        '한국경제(2026.5·7)·머니투데이(2026.8)·토요경제(2026.11)·뉴데일리(2021.11·2026.1)·유니콘팩토리(2026.3)·아시아경제(2025.5). 회사 묶음은 당사 정리')

# (회사, 업종, 본업, [(시점, 넓힌 영역, 한 줄 설명)], 큰 숫자, 숫자 설명)
ROWS = [
    ('오늘의집', '인테리어 플랫폼', '집들이 커뮤니티·커머스',
     [("'19", '시공 중개', '시공업체 견적 연결'), ("'22.1", '이사 견적', '이사 업체 연결'),
      ("'25.9", '직영 부분시공', '시공 자회사 출범'), ("'26.10", '무빙패스', '이사 고객 전용 멤버십')],
     '약 3배', '직영 부분시공 계약액\n(2026.6)\n전년 같은 달 대비'),
    ('한샘', '가구 대기업', '가구·부엌 판매',
     [("'19~", '리하우스 시공', '대리점 86→682곳'), ("'25.9", 'A/S 직영', '자회사 A/S를 본사로'),
      ("'26.6", '인테리어 플래너', '상담~시공~AS 한 앱'), ("'26.7", '홈닉 제휴', '단지 평면 3D 인테리어')],
     '+129억 원', '리하우스 영업이익\n(2026년 1~3분기)\n전년 −31억 → 흑자'),
    ('삼성물산 홈닉', '건설사 플랫폼', '래미안 아파트 시공',
     [("'23.8", '입주민 앱 홈닉', '래미안 원베일리 첫 도입'), ("'24.12", '타사 단지 도입', '한화 건설부문부터'),
      ("'26.5", '입주 전 서비스', '인테리어·청소·세금'), ("'26.8", '입주 후 생활', '청소·세차·펫케어')],
     '21곳', '홈닉 협약 건설사\n(2026.11)\n7월 18곳에서 늘어남'),
    ('LG전자', '가전', '가전 제조·판매',
     [("'23.7", '가전 구독', '정기 관리·A/S 포함'), ("'24.8", 'ThinQ 3D 홈뷰', '우리 집 평면도로 관리'),
      ("'24.12", 'ThinQ API 개방', '외부 서비스와 연결')],
     '2조 원+', '가전 구독 매출\n(2025년)\n실적 발표 기준'),
    ('직방', '부동산 앱', '원룸·아파트 매물 정보',
     [("'18.4", '시세 정보', '호갱노노 인수'), ("'21.6", '아파트 중개', '자회사 중개법인'),
      ("'22.7", '스마트홈 기기', '삼성SDS 홈IoT 인수')],
     '379억 원', '스마트홈 중심 상품매출\n(2025년)\n전년 대비 −15%'),
    ('무브제', '이사 컨시어지', '이사(통인익스프레스, 1972~)',
     [("'19", '이주 컨시어지', '3D 배치도·예약 조율'), ('시점 미공개', '부동산 컨시어지', '매물·계약·전입 신고'),
      ('시점 미공개', '홈케어 컨시어지', '입주청소·정리·배관')],
     '3개', '컨시어지 영역\n이사·부동산·홈케어\n(회사 소개)'),
]

X0, X1 = 0.5625, 9.4349
CO_W = 1.62                 # 회사·본업 칸
AR_X = X0 + CO_W + 0.05     # 화살표
CH_X = AR_X + 0.27          # 칩 시작
NUM_W = 2.0
NUM_X = X1 - NUM_W
CH_GAP = 0.06
CH_W = (NUM_X - 0.12 - CH_X - 3 * CH_GAP) / 4
HY, RY, RH, RG = 1.6, 1.84, 0.47, 0.04

# 머리 줄
text(s, X0, HY, CO_W, 0.18, '회사 · 원래 본업', 7, True, MUTE, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
text(s, CH_X, HY, 4 * CH_W + 3 * CH_GAP, 0.18, '넓힌 영역 (시점순, 파랑 = 가장 최근)', 7, True, MUTE, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
text(s, NUM_X, HY, NUM_W, 0.18, '확인된 숫자', 7, True, MUTE, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)

for k, (co, cat, base, chips, num, cap) in enumerate(ROWS):
    y = RY + k * (RH + RG)
    # 회사·본업 (회색)
    box(s, X0, y, CO_W, RH, LIGHT)
    rich(s, X0 + 0.1, y, CO_W - 0.14, RH,
         [(cat, 6, True, ACC), (co, 9.5, True, INK), (base, 6.5, False, SUB)])
    # 파랑 화살표
    box(s, AR_X + 0.03, y + RH / 2 - 0.08, 0.17, 0.16, ACC, MSO_SHAPE.CHEVRON)
    # 넓힌 영역 칩
    for j, (when, area, det) in enumerate(chips):
        cx = CH_X + j * (CH_W + CH_GAP)
        last = j == len(chips) - 1
        box(s, cx, y, CH_W, RH, ACC if last else LAV)
        rich(s, cx + 0.08, y, CH_W - 0.12, RH,
             [(when, 6, True, PALE if last else ACC), (area, 8, True, WHITE if last else INK),
              (det, 6, False, PALE if last else SUB)])
    # 확인된 숫자
    box(s, NUM_X, y, NUM_W, RH, LIGHT)
    text(s, NUM_X + 0.1, y, 0.95, RH, num, 13, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    text(s, NUM_X + 1.06, y, NUM_W - 1.1, RH, cap, 6, False, SUB, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)

prs.save(OUT); print('saved', OUT, len(prs.slides))
