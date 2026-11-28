"""2026-11-28 IR 덱 MARKET 장 1장: 주거 산업이 흘러가는 방향을 키워드 4개로(덱 사본 맨 끝에 추가).
  01 집의 데이터화 / 02 사는 것에서 맡기는 것으로 / 03 로봇 친화 설계 / 04 전통 강자의 경계 넘기
숫자·출처: 리서치/확장/주거 산업 키워드 흐름 2026-11-28.md (DART 감사보고서, 회사 발표, 국토부 발표 보도, 언론 보도)
톤: build_ir_feedback_slides.py와 같은 헬퍼·색 값, OVERVIEW 장의 CONTENT 레이아웃·자리표시자 서식 복사
사용: python build_ir_industry_keywords.py <덱.pptx> <출력.pptx>"""
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

s = prs.slides.add_slide(REF.slide_layout)
fill_ph(s, 'MARKET', '물건 판매에서 집을 이해하고 맡아 주는 쪽으로 가는 주거 산업',
        '집은 데이터가 되고, 서비스는 맡기는 쪽으로, 건물은 로봇을 들이고, 전통 강자는 업종 경계를 넘는 흐름',
        '출처: 삼성전자 뉴스룸(2024.12), LG전자 공지(2024.8), 숨고 감사보고서(DART, 별도), LG전자 실적 발표 보도(2026.1), 청소연구소 회사 발표(2026.1 보도), 국토부 발표 보도(2025.5), 아이뉴스24(2026.9), '
        '한샘 실적 보도(CEO스코어데일리 2026.11)·쿠키뉴스(2026.6)·한국경제(2026.7), 비즈워치(2022.7). 과거·지금 문구는 당사 정리')

# (번호, 키워드, 과거, 지금, 큰 숫자, 숫자 설명, 사례[(회사, 한 줄)])
COLS = [
    ('01', '집의 데이터화',
     '도면은 종이, 가전은 따로', '평면도 위에 가전·기록이\n앱 데이터로 쌓임',
     '2,000만 명', '삼성 SmartThings 국내 이용 고객\n(2024.12, 계정 연결 등록 기준)',
     [('삼성 SmartThings', '로봇청소기로 평면도'), ('LG 씽큐 3D 홈뷰', '우리 집 평면도'), ('GS건설 자이홈', '점검·AS·스마트홈')]),
    ('02', '사는 것 → 맡기는 것',
     '물건을 사서 직접 처리', '사람·관리를 맡기고\n매달 돈을 냄',
     '4.7배', '숨고 매출 152억 → 720억 원\n(2021 → 2025, 감사보고서)',
     [('숨고', '이사·청소·시공 연결'), ('LG전자 가전 구독', '2025년 2조 원+'), ('청소연구소', '매니저 9만→20만')]),
    ('03', '로봇 친화 설계',
     '사람만 다니는 단지', '로봇이 공동현관·승강기를\n지나 문 앞까지',
     '180억 원', '국토부 로봇 친화형 건축물\n기술개발 정부 예산(2025~2028)',
     [('현대건설', '디에이치 대치 배송 로봇'), ('삼성물산', '래미안 배달 로봇 상용화'), ('국토부', '기술표준·인증 개선 추진')]),
    ('04', '전통 강자의 경계 넘기',
     '가구는 팔고, 집은 지으면 끝', '시공·플랫폼·제휴로\n주거 공간 전체로',
     '+129억 원', '한샘 리하우스(리모델링) 영업이익\n2026 1~3분기, 전년 −31억에서 흑자',
     [('한샘', 'A/S 직영·시공 앱·홈닉 제휴'), ('삼성물산 홈닉', '건설사→입주 플랫폼'), ('직방', '중개→스마트홈 인수')]),
]
X0, CW, GAP = 0.5625, 2.09, 0.17
for k, (no, kw, past, now, big, bigdesc, cases) in enumerate(COLS):
    x = X0 + k * (CW + GAP)
    # 머리: 파랑 원 번호 + 키워드
    c = box(s, x, 1.6, 0.3, 0.3, ACC, MSO_SHAPE.OVAL)
    text(s, x, 1.6, 0.3, 0.3, no, 8, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(s, x + 0.38, 1.6, CW - 0.38, 0.3, kw, 10.5, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    # 과거 → 지금
    box(s, x, 2.0, CW, 0.28, LIGHT)
    text(s, x + 0.1, 2.0, 0.4, 0.28, '과거', 6.5, True, MUTE, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    text(s, x + 0.45, 2.0, CW - 0.5, 0.28, past, 7.5, False, SUB, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    ar = box(s, x + CW / 2 - 0.09, 2.31, 0.18, 0.11, ACC, MSO_SHAPE.ISOSCELES_TRIANGLE); ar.rotation = 180
    box(s, x, 2.45, CW, 0.44, LAV)
    text(s, x + 0.1, 2.45, 0.4, 0.44, '지금', 6.5, True, ACC, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    text(s, x + 0.45, 2.45, CW - 0.5, 0.44, now, 7.5, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    # 큰 숫자
    text(s, x, 3.0, CW, 0.42, big, 24, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    text(s, x, 3.44, CW, 0.3, bigdesc, 6.5, False, SUB)
    hline(s, x, 3.82, x + CW)
    # 대표 사례 칩
    text(s, x, 3.88, CW, 0.16, '대표 사례', 6.5, True, MUTE)
    for j, (nm, d) in enumerate(cases):
        cy = 4.07 + j * 0.27
        box(s, x, cy, CW, 0.24, LIGHT)
        text(s, x + 0.08, cy, CW - 0.12, 0.24, nm, 7, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
        text(s, x + 0.08, cy, CW - 0.14, 0.24, d, 6, False, SUB, PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
prs.save(OUT); print('saved', OUT, len(prs.slides))
