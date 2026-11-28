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
ACC2 = RGBColor(0x8E, 0x8A, 0xFF)
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
        '출처: 삼성전자 발표(SmartThings 이용자 2020·2023.10·2024.9·2025.12, 국내 2,000만 2024.12), 숨고 감사보고서(DART), LG전자 발표 보도(2026.1), 로봇 단계는 뉴시스·경찰청·바이라인네트워크·국토부 보도·아이뉴스24(2022~2026), 한샘 실적 보도(CEO스코어데일리 2026.11). 과거·지금 문구는 당사 정리')

# (번호, 키워드, 과거, 지금, 헤드라인, 그래프 설명, 그래프, 사례)
# 그래프: ('bar', 라벨, 값, 값표시) / ('step', 라벨, 설명)
COLS = [
    ('01', '집의 데이터화', '도면은 종이, 가전은 따로', '평면도 위에 가전·기록이\n앱 데이터로 쌓임',
     '5년 새 3.6배', '삼성 SmartThings 전 세계 이용자(억 명, 등록 기준)',
     ('bar', ["'20", "'23.10", "'24.9", "'25.12"], [1.2, 2.9, 3.5, 4.3], ['1.2', '2.9', '3.5', '4.3'], False),
     [('삼성 SmartThings', '국내 이용 고객 2,000만'), ('LG 씽큐 3D 홈뷰', '우리 집 평면도'), ('GS건설 자이홈', '점검·AS·스마트홈')]),
    ('02', '사는 것 → 맡기는 것', '물건을 사서 직접 처리', '사람·관리를 맡기고\n매달 돈을 냄',
     '4년간 4.7배', '숨고 매출(억 원, 감사보고서)',
     ('bar', ['21', '22', '23', '24', '25'], [151.8, 313.8, 458.6, 624.2, 719.8], ['152', '314', '459', '624', '720'], False),
     [('숨고', '이사·청소·시공 연결'), ('LG전자 가전 구독', '2025년 2조 원+'), ('청소연구소', '매니저 9만→20만')]),
    ('03', '로봇 친화 설계', '사람만 다니는 단지', '로봇이 공동현관·승강기를\n지나 문 앞까지',
     '4년 만에 실증→상용화', '로봇이 집 앞까지 오기까지',
     ('step', ["'22", "'23", "'24", "'25", "'26"], ['사설\n인증', '보도\n통행', '첫\n안전인증', '국토부\n기술개발', '단지\n상용화']),
     [('현대건설', '디에이치 대치 배송 로봇'), ('삼성물산', '래미안 배달 로봇 상용화'), ('국토부', '기술표준·인증 개선 추진')]),
    ('04', '전통 강자의 경계 넘기', '가구는 팔고, 집은 지으면 끝', '시공·플랫폼·제휴로\n주거 공간 전체로',
     '리모델링 흑자 전환', '한샘 리하우스 영업이익(억 원, 1~3분기)',
     ('bar', ['2025', '2026'], [-31, 129], ['2025  −31', '2026  +129'], True),
     [('한샘', 'A/S 직영·시공 앱·홈닉 제휴'), ('삼성물산 홈닉', '건설사→입주 플랫폼'), ('직방', '중개→스마트홈 인수')]),
]
GA_Y, GA_H = 3.34, 0.36   # 그래프 막대 영역
def minibar(x, w, cats, vals, labs, signed):
    n = len(vals); slot = w / n; bw = slot * 0.55
    vmax = max(vals); vmin = min(0, min(vals)); rng = vmax - vmin
    zero = GA_Y + GA_H * (vmax / rng)
    if signed: hline(s, x, zero, x + w, LINE)
    for i, (c, v, lb) in enumerate(zip(cats, vals, labs)):
        bx = x + i * slot + (slot - bw) / 2
        hgt = max(GA_H * abs(v) / rng, 0.02)
        top = zero - hgt if v >= 0 else zero
        last = i == n - 1
        box(s, bx, top, bw, hgt, (ACC if last else ACC2) if v >= 0 else RGBColor(0xC4, 0xC4, 0xCA))
        ly = top - 0.15 if v >= 0 else top + hgt + 0.01
        text(s, bx - 0.3, ly, bw + 0.6, 0.14, lb, 6.5, last, INK if last else SUB, PP_ALIGN.CENTER)
        if not signed: text(s, bx - 0.15, GA_Y + GA_H + 0.02, bw + 0.3, 0.13, c, 6, False, MUTE, PP_ALIGN.CENTER)
def ministep(x, w, cats, labs):
    n = len(cats); slot = w / n; bw = slot * 0.82
    for i, (c, lb) in enumerate(zip(cats, labs)):
        bx = x + i * slot + (slot - bw) / 2
        hgt = 0.08 + i * 0.06; top = GA_Y + GA_H - hgt
        last = i == n - 1
        box(s, bx, top, bw, hgt, ACC if last else (ACC2 if i >= 2 else LAV))
        text(s, bx - 0.05, top - 0.25, bw + 0.1, 0.24, lb, 5.5, last, INK if last else SUB, PP_ALIGN.CENTER, MSO_ANCHOR.BOTTOM)
        text(s, bx, GA_Y + GA_H + 0.02, bw, 0.13, c, 6, False, MUTE, PP_ALIGN.CENTER)
X0, CW, GAP = 0.5625, 2.09, 0.17
for k, (no, kw, past, now, head, cap, g, cases) in enumerate(COLS):
    x = X0 + k * (CW + GAP)
    box(s, x, 1.6, 0.3, 0.3, ACC, MSO_SHAPE.OVAL)
    text(s, x, 1.6, 0.3, 0.3, no, 8, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(s, x + 0.38, 1.6, CW - 0.38, 0.3, kw, 10.5, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    box(s, x, 1.98, CW, 0.24, LIGHT)
    text(s, x + 0.1, 1.98, 0.4, 0.24, '과거', 6.5, True, MUTE, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    text(s, x + 0.45, 1.98, CW - 0.5, 0.24, past, 7.5, False, SUB, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    ar = box(s, x + CW / 2 - 0.08, 2.24, 0.16, 0.09, ACC, MSO_SHAPE.ISOSCELES_TRIANGLE); ar.rotation = 180
    box(s, x, 2.35, CW, 0.4, LAV)
    text(s, x + 0.1, 2.35, 0.4, 0.4, '지금', 6.5, True, ACC, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    text(s, x + 0.45, 2.35, CW - 0.5, 0.4, now, 7.5, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    # 그래프
    text(s, x, 2.82, CW, 0.24, head, 12, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
    text(s, x, 3.05, CW, 0.14, cap, 6, False, SUB)
    if g[0] == 'bar': minibar(x + 0.05, CW - 0.1, g[1], g[2], g[3], g[4])
    else: ministep(x + 0.02, CW - 0.04, g[1], g[2])
    hline(s, x, 3.9, x + CW)
    text(s, x, 3.94, CW, 0.15, '대표 사례', 6.5, True, MUTE)
    for j, (nm, d) in enumerate(cases):
        cy = 4.11 + j * 0.25
        box(s, x, cy, CW, 0.22, LIGHT)
        text(s, x + 0.08, cy, CW - 0.12, 0.22, nm, 7, True, INK, PP_ALIGN.LEFT, MSO_ANCHOR.MIDDLE)
        text(s, x + 0.08, cy, CW - 0.14, 0.22, d, 6, False, SUB, PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
prs.save(OUT); print('saved', OUT, len(prs.slides))
