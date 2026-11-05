"""Render the 그날의 류 full report in a purpose-built analyst layout.

Content blocks come from build_report_v3.py (the section between the Investment Highlights
marker and doc.save); this file only decides how they look.
"""
import re
import sys
from pathlib import Path

import docx
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.shared import Pt, Cm, RGBColor

CONTENT, CHARTS, OUT = Path(sys.argv[1]), Path(sys.argv[2]), sys.argv[3]

FONT = '맑은 고딕'
NAVY, ACCENT, INK, MUTED = '1F3A5F', '2A78D6', '1A1A1A', '6B7280'
RULE, ZEBRA, PANEL = 'C9D2DE', 'F5F7FA', 'EEF3F9'
TEXT_W_CM = 17.0
TWIP = 567  # per cm

doc = docx.Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(2.0)
sec.top_margin, sec.bottom_margin = Cm(2.3), Cm(2.0)
sec.header_distance, sec.footer_distance = Cm(1.1), Cm(1.0)
sec.different_first_page_header_footer = True


def rgb(h):
    return RGBColor.from_string(h)


def style_font(st, size, bold=False, color=INK):
    st.font.name = FONT
    st.element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), FONT)
    st.font.size = Pt(size); st.font.bold = bold; st.font.color.rgb = rgb(color)


normal = doc.styles['Normal']
style_font(normal, 9.5)
normal.paragraph_format.line_spacing = 1.32
normal.paragraph_format.space_after = Pt(3)
for name, size, before, after in (('Heading 1', 15, 0, 8), ('Heading 2', 12, 14, 6), ('Heading 3', 10.5, 10, 4)):
    st = doc.styles[name]
    style_font(st, size, True, NAVY if name != 'Heading 3' else INK)
    pf = st.paragraph_format
    pf.space_before, pf.space_after, pf.keep_with_next = Pt(before), Pt(after), True
    pf.line_spacing = 1.15


def shade(cell, fill):
    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill}"/>'))


def borders(el, **sides):
    """sides: top/bottom/left/right/insideH/insideV = (size_eighths, color) or None (no line)."""
    tag = 'w:tblBorders' if el.tag == qn('w:tblPr') else 'w:tcBorders'
    b = el.find(qn(tag))
    if b is None:
        b = parse_xml(f'<{tag} {nsdecls("w")}/>'); el.append(b)
    for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        if side not in sides:
            continue
        v = sides[side]
        xml = (f'<w:{side} {nsdecls("w")} w:val="nil"/>' if v is None else
               f'<w:{side} {nsdecls("w")} w:val="single" w:sz="{v[0]}" w:space="0" w:color="{v[1]}"/>')
        old = b.find(qn(f'w:{side}'))
        if old is not None:
            b.remove(old)
        b.append(parse_xml(xml))


def para_border(p, side, size, color, space=4):
    ppr = p._p.get_or_add_pPr()
    pbdr = ppr.find(qn('w:pBdr'))
    if pbdr is None:
        pbdr = parse_xml(f'<w:pBdr {nsdecls("w")}/>'); ppr.append(pbdr)
    pbdr.append(parse_xml(f'<w:{side} {nsdecls("w")} w:val="single" w:sz="{size}" w:space="{space}" w:color="{color}"/>'))


def runs(p, text, size=None, color=None, bold=None):
    text = re.sub(r'\[\[([^\]|]+)\|([^\]]+)\]\]', r'\2', text)
    for part in re.split(r'(\*\*[^*]+\*\*)', text):
        if not part:
            continue
        b = part.startswith('**')
        r = p.add_run(part[2:-2] if b else part)
        if b or bold:
            r.bold = True
        if size:
            r.font.size = Pt(size)
        if color:
            r.font.color.rgb = rgb(color)
    return p


def fit(widths):
    tw = int(TEXT_W_CM * TWIP)
    ws = [int(tw * w / sum(widths)) for w in widths]
    ws[-1] += tw - sum(ws)
    return ws


def fix_table(t, ws):
    tblpr = t._tbl.tblPr
    for tag in ('w:tblW', 'w:tblLayout', 'w:tblInd'):
        for x in tblpr.findall(qn(tag)):
            tblpr.remove(x)
    tblpr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{sum(ws)}" w:type="dxa"/>'))
    tblpr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    tblpr.append(parse_xml(f'<w:tblCellMar {nsdecls("w")}><w:top w:w="45" w:type="dxa"/><w:left w:w="85" w:type="dxa"/>'
                           f'<w:bottom w:w="45" w:type="dxa"/><w:right w:w="85" w:type="dxa"/></w:tblCellMar>'))
    grid = t._tbl.tblGrid
    for i, g in enumerate(grid.findall(qn('w:gridCol'))):
        g.set(qn('w:w'), str(ws[i]))
    for row in t.rows:
        for i, c in enumerate(row.cells):
            c.width = ws[i]
            tcw = c._tc.get_or_add_tcPr().find(qn('w:tcW'))
            tcw.set(qn('w:type'), 'dxa'); tcw.set(qn('w:w'), str(ws[i]))


def cell_text(c, text, size=8.5, bold=False, color=INK, align=WD_ALIGN_PARAGRAPH.LEFT):
    c.text = ''
    lines = str(text).split('\n')
    for k, ln in enumerate(lines):
        p = c.paragraphs[0] if k == 0 else c.add_paragraph()
        p.alignment = align
        pf = p.paragraph_format
        pf.space_after, pf.space_before, pf.line_spacing = Pt(0), Pt(0), 1.15
        runs(p, ln, size=size, color=color, bold=bold)
    c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def keep_next(p):
    p.paragraph_format.keep_with_next = True


COUNT = {'표': 0, '그림': 0}


def caption(kind, title):
    COUNT[kind] += 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(8), Pt(3)
    keep_next(p)
    runs(p, f'[{kind} {COUNT[kind]}] ', size=9, color=ACCENT, bold=True)
    runs(p, title, size=9, color=NAVY, bold=True)


def source(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(2), Pt(8)
    runs(p, text, size=7.5, color=MUTED)


# ---------- content primitives (names used by the content script) ----------
H = {'h1': 0, 'h2': 0}
SUMMARY = {
    '시장 분석': '무대는 소형 디바이스이며, 시장은 합산해 넓히지 않고 같은 고객의 다음 문제를 따라 단계별로 만나는 시장을 나란히 적음. 출발점은 생리 지연 고민(25~39세 주기 불규칙 응답 약 155만~160만 명, D)이고, 규모는 첫 제품 기준 TAM 약 406만 명(D, 보정 시 378만~403만) → SAM(비약물 셀프케어 시도자) 약 69만~166만 명 → SOM(첫해 구매자) 140~4,440명(D)으로 좁힘. 국내 7사의 생리 전후 표방 SKU는 0개(C)이나 쿠팡에는 생리 표방 저가 기기 10개(중앙값 2.6만 원, A)가 있고 생리통 검색 상위 기기는 3만 원대(A). 확장 선례는 있으나 기기 → 기기 확장 기업은 매출이 줄었음(세라젬 −27%, 풀리오 −29%, A). 7개 세부 시장 비교에서 생리 전후는 수면과 공동 1위(동일 가중, D)이나 가중치에 따라 1~6위이며, 국내 소형 마사지기는 약 4,000억~9,000억 원(D). 효능은 약속하지 않으며 효능을 말하려면 의료기기·임상 경로를 따로 선택해야 함.',
    '사업 분석 및 투자포인트': '생리 지연 고민에서 출발한 발목 지압 밴드가 첫 검증 가설. 구매 목적을 셋으로 나눠 확인하고, 주기 변화는 근거(지압 주기 임상 0건)·표현 규제·임신 가능성 때문에 약속하지 않음. 복부 온열은 수요는 크나 3.5만 원대 전동·1,200원 패치가 촘촘해, 같은 고객이 기존 제품을 그만둔 이유와 더 높은 지불 의향이 확인될 때만 시험. 문제·지불 이유가 약하면 시제품 이전에도 멈춤. 투자포인트는 해자가 아니라 시행착오 기록으로 역량을 쌓는 구조.',
    '경쟁 분석': '생리통 착용형(Beurer·Myoovi·Livia·Ova+, 미국 적응증은 진통제 병용 조건부), 복부 온열 착용(Stoov), 공통 본체·스트랩 구조(Hyperice·Therabody)가 이미 있어 형태·공용화는 해자가 아님. 고객이 생리통 검색에서 처음 보는 가격은 쿠팡 상위 전동 기기 중앙값 34,900원(A)이며 국내 7사 제품은 그 상위 91개에 없음. 7사는 매출의 15~38%를 광고에 쓰는 온라인 D2C이고, 풀리오 매출 −29%의 원인은 공시에 없음. 에이피알은 전용 소모품이 아니라 기기와 화장품의 병용·교차판매 모델.',
    '재무계획': '기기 단독 손익분기는 약 6,286대로 3년차 계획 5,000대에서는 -0.45억 원(린캔버스 팀 가정, 산정 근거 미기재, D. CAC 1.5만 원·환불 0% 전제이며 CAC는 v4 기본 3.6만 원과 미정합, 재무 반영 시 확정). 패치는 교체·재구매가 확인되기 전 기본 시나리오에서 제외. 흑자 경로는 가격·원가·판매량·후속 제품 등 여러 가지. 모든 손익은 계획 시나리오.',
    '매출추정 및 사업가치': '첫해 판매 규모는 시장 깔때기로 구매자 약 140~4,440명(기본 960명), 첫해 매출 약 0.12억~5.73억 원(소비자가, 계산 D, 방문 수 미측정). 방문 1만·4만·12만 명은 펀딩 1회와 이후 자사몰 11개월 방문을 더한 대응점으로 설명함. 규모 예시로 심한 월경통 응답 여성 약 406만 명(추정)의 1%가 누적으로 본체를 사면 약 49억 원(연간 아님, 계산: 406만 × 1% × 판매가 12만 원(부가세 제외), D. 1%는 설명용 가정·근거 미확보, 12만 원은 린캔버스 팀 가정·산정 근거 미기재). 외부 근거 감사(2026-11-05)는 전환율 첫해 보정, 가격 축 분리, 침투 분모 SAM·3년, 후속 구매·패치 재구매 하향 대안을 제시했으며 재무 반영 시 확정, 본 판 수치는 유지. 실패 원인별 전환 경로를 정하고, 기록은 건수가 아니라 다음 제품에서의 재사용 결과로만 역량 축적으로 셈.',
    '리스크 및 대응': '가장 큰 리스크는 차별화(이미 있는 착용형 제품 대비 선택 이유), 효능 표현·규제, 수요 미검증, 명칭 충돌(상표). 안전은 생리 지연 고객의 임신 여부 확인 안내와 임신 중 사용 금지, 온열 저온화상 대비 표면온도 상한·타이머·수면 중 사용 금지를 설계 조건으로 둠. 국내 7사처럼 단일 히트 제품에 매출이 몰리는 구조도 리스크로 봄. 확인은 날짜가 아니라 단계(100일 검증·확정 견적·식약처 질의·변리사·한의사 자문)로 둠.',
    '결론': '첫 제품은 생리 지연 고민에서 출발한 발목 지압 밴드이며, 하던 일을 내려놓지 않아도 되는 그날의 케어로 진입을 검증. 확장은 같은 고객의 다음 문제를 따라 생리 전후 케어 → 바디 웰니스 → 홈헬스·홈뷰티(웨어러블 연동은 곁가지)로 넓힌다는 가설이며, 단계 조건을 넘을 때만 진행. 고객·제품·실행 세 기준으로 진행과 중단을 가름.',
}


def callout(text, label='핵심 요약'):
    t = doc.add_table(rows=1, cols=1)
    fix_table(t, [int(TEXT_W_CM * TWIP)])
    c = t.cell(0, 0); shade(c, PANEL)
    borders(t._tbl.tblPr, top=None, bottom=None, left=(24, ACCENT), right=None, insideH=None, insideV=None)
    c.text = ''
    p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(2)
    runs(p, label, size=8.5, color=ACCENT, bold=True)
    p2 = c.add_paragraph(); p2.paragraph_format.space_after = Pt(1); p2.paragraph_format.line_spacing = 1.3
    runs(p2, text, size=9.5, color=INK)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def h1(t):
    H['h1'] += 1; H['h2'] = 0
    p = doc.add_paragraph(style='Heading 1')
    p.paragraph_format.page_break_before = True  # every chapter opens a page (chapter 1 follows the TOC)
    runs(p, f'{H["h1"]}. {t}')
    para_border(p, 'bottom', 12, NAVY, 4)
    if t == 'Investment Highlights':
        overview()
    if t in SUMMARY:
        callout(SUMMARY[t])


def h2(t):
    H['h2'] += 1
    p = doc.add_paragraph(style='Heading 2')
    runs(p, f'{H["h1"]}.{H["h2"]}  {t}')


def s1(t):
    # topic heading (level 3 in the TOC): accent bar instead of a ■ glyph so the TOC text stays clean
    p = doc.add_paragraph(style='Heading 3')
    para_border(p, 'left', 18, ACCENT, 6)
    runs(p, t)


def bullet(text, level):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    indent = 0.45 + 0.55 * level
    pf.left_indent, pf.first_line_indent = Cm(indent), Cm(-0.35)
    pf.space_after = Pt(3)
    runs(p, ('•' if level == 0 else '–') + '  ', color=ACCENT if level == 0 else MUTED)
    runs(p, text)


def s2(*ts):
    for t in ts:
        bullet(t, 0)


def s3(*ts):
    for t in ts:
        bullet(t, 1)


def blank():
    pass


def add(kind, text=''):
    {'s1': s1, 's2': s2, 's3': s3, 'h1': h1, 'h2': h2}[kind](text)


def cap(title):
    return title


def table(title, header, rows, src, widths=None):
    caption('표', re.sub(r'^\[도표 \d+\] ', '', title))
    n = len(header)
    ws = fit(widths or [1] * n)
    prose = [any(len(str(r[i])) > 16 for r in rows) for i in range(n)]
    t = doc.add_table(rows=1 + len(rows), cols=n)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    fix_table(t, ws)
    borders(t._tbl.tblPr, top=(12, NAVY), bottom=(12, NAVY), insideH=(4, RULE), left=None, right=None, insideV=None)
    hdr = t.rows[0]
    hdr._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    for i, h in enumerate(header):
        shade(hdr.cells[i], NAVY)
        cell_text(hdr.cells[i], h, bold=True, color='FFFFFF', align=WD_ALIGN_PARAGRAPH.CENTER)
    for r, row in enumerate(rows, start=1):
        tr = t.rows[r]
        tr._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for i in range(n):
            v = row[i] if i < len(row) else ''
            c = tr.cells[i]
            if r % 2 == 0:
                shade(c, ZEBRA)
            align = WD_ALIGN_PARAGRAPH.LEFT if (i == 0 or prose[i]) else WD_ALIGN_PARAGRAPH.CENTER
            if i == 0 and not prose[0]:
                align = WD_ALIGN_PARAGRAPH.CENTER
            cell_text(c, v, bold=(i == 0), color=NAVY if i == 0 else INK, align=align)
    if len(rows) <= 8:  # short tables stay on one page with their caption
        for tr in t.rows[:-1]:
            for c in tr.cells:
                for p in c.paragraphs:
                    keep_next(p)
    source(src)


def figure(title, png, src, width_cm=15.0):
    caption('그림', re.sub(r'^\[도표 \d+\] ', '', title))
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    keep_next(p)
    p.add_run().add_picture(str(png), width=Cm(width_cm))
    source(src)


def kv_table(pairs, label_w=3.3):
    """Two pairs per row: label | value | label | value."""
    rows = [pairs[i:i + 2] for i in range(0, len(pairs), 2)]
    t = doc.add_table(rows=len(rows), cols=4)
    vw = (TEXT_W_CM - 2 * label_w) / 2
    ws = fit([label_w, vw, label_w, vw])
    fix_table(t, ws)
    borders(t._tbl.tblPr, top=(12, NAVY), bottom=(12, NAVY), insideH=(4, RULE), left=None, right=None, insideV=None)
    for r, pr in enumerate(rows):
        cells = t.rows[r].cells
        if len(pr) == 1:  # last odd pair spans the row
            a = cells[1].merge(cells[3]); pr = [pr[0]]
        for k, (lab, val) in enumerate(pr):
            shade(cells[2 * k], PANEL)
            cell_text(cells[2 * k], lab, bold=True, color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER)
            cell_text(cells[2 * k + 1], val)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def overview():
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(4)
    runs(p, '핵심 지표', size=10.5, color=INK, bold=True)  # not a heading: kept out of the 3-level TOC (no number before 1.1)
    kv_table([('주기 불규칙 응답 (25~39세)', '약 155만~160만 명 (조사 34.9~36.2% × 인구, D. 출발 문제)'), ('생리불순 진료 (N91)', '50만 4,722명 (2025, 심평원, A. 진료 인원, 406만과 합산 안 함)'),
              ('심한 월경통 응답', '성인 여성 41.6% (현재 월경 중인 성인 기준, A)'), ('산부인과·비뇨의학과 진료 경험', '심한 증상 응답자 중 28.5% (A)'),
              ('약국 진통제 대처', '59.2% (심한 증상 성인 1,266명 중, 비정상 자궁출혈 포함, A)'), ('TAM: 심한 월경통 응답 여성', '약 406만 명 (19~49세, D. 월경 상태 보정 시 378만~403만)'), ('SAM: 비약물 셀프케어 시도자', '약 69만~166만 명 (중앙 122만, TAM × 17~41%, D)'), ('SOM: 첫해 구매자', '140~4,440명, 기본 960명 (0.12억~5.73억 원, D)'),
              ('소형 안마기 공시 4사 매출', '1,930억 원 (2025, A 합산, 비공시 제외 하한)'), ('(확장 무대 참고) 국내 홈뷰티 디바이스', '1.6조 원(2022, 삼일PwC 원문 확인·정의 미기재) → 3.4조 원(2030E, 전망 주체 미표기) (B)'),
              ('W케어 상품 수', '약 1만 8천 개, 2022년 대비 약 5배 (B)'), ('1% 침투 예시', '누적 본체 매출 약 49억 원 (계산: 406만 × 1% × 판매가 12만 원, D. 1%는 근거 없는 설명용 가정, 12만 원은 린캔버스 팀 가정)')])


# ---------- header / footer ----------
def page_field(p):
    r = p.add_run()
    r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>'))
    r2 = p.add_run(); r2._r.append(parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>'))
    r3 = p.add_run(); r3._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>'))
    for x in (r, r2, r3):
        x.font.size = Pt(8); x.font.color.rgb = rgb(MUTED)


hp = sec.header.paragraphs[0]
hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
runs(hp, '그날의 류  |  산업·기업 분석 보고서  |  2026.11 (v10.11)', size=8, color=MUTED)
para_border(hp, 'bottom', 4, RULE, 3)
fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
page_field(fp)

# ---------- cover ----------
for _ in range(5):
    doc.add_paragraph()
p = doc.add_paragraph(); runs(p, 'INDUSTRY & COMPANY REPORT  |  사내벤처 분석 보고서', size=10, color=ACCENT, bold=True)
p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(6); p.paragraph_format.line_spacing = 1.0
runs(p, '그날의 류', size=34, color=NAVY, bold=True)
p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(14)
runs(p, '하던 일을 내려놓지 않아도 되는 그날의 케어 — 생리 지연 고민에서 출발한 발목 밴드로 시작해 소형 헬스·뷰티 디바이스로 넓혀 가려는 브랜드', size=13, color=INK)
para_border(p, 'bottom', 18, NAVY, 10)
kv_table([('작성일', '2026년 11월 5일 (v10.11)'), ('형식', '산업분석 + 기업분석 풀리포트')])
callout('그날의 류는 "생리 지압기"가 아니라, 여성이 반복되는 일상 불편 때문에 하던 일을 내려놓지 않도록 착용형 케어 제품을 만드는 '
        '브랜드다. 첫 제품은 생리 지연 고민에서 출발한 발목 지압 밴드이며, 같은 고객·채널·역량으로 생리 전후 케어 → 바디 웰니스 → 소형 홈헬스·홈뷰티(웨어러블 연동은 곁가지)로 넓힐 수 있다는 가설을 '
        '단계마다 조건으로 확인한다(같은 고객의 교차 구매는 공개 자료가 없어 직접 잰다). 효능을 약속하지 않고 고객의 상황으로 설명하며, 조건을 채우지 못하면 고치거나 바꾸거나 멈춘다. 다음 단계에 다시 쓰일 고객·시험 기록을 남기는 것을 목표로 한다.',
        label='한 문장 논지')
for _ in range(4):
    doc.add_paragraph()
p = doc.add_paragraph(); runs(p, 'Compliance Notice', size=8, color=NAVY, bold=True)
p = doc.add_paragraph()
runs(p, '본 자료는 사내벤처 심사를 위한 내부 검토 자료로, 팀 지식 저장소(teoo)의 근거에 기반해 작성되었습니다. 수치 뒤의 (A)~(D)는 근거 등급이며 '
        'A 공시·정부 통계·규제 문서·논문 원문, B 리포트·언론 원문 인용, C 업체 주장·판매용 요약·AI 답변, D 팀 가설·내부 추정을 뜻합니다. '
        '추정·시나리오는 실적이나 검증 결과가 아니며, 제품의 의학적 효능을 주장하지 않습니다.', size=7.5, color=MUTED)

# ---------- v8 대비 변경 요약 (one page after the cover; data at the top of report_content.py) ----------
sys.path.insert(0, str(Path(__file__).resolve().parent))
from changes import load_changes  # noqa: E402

CH = load_changes(CONTENT.read_text(encoding='utf-8'))
if CH['rows']:
    assert f'({CH["latest"]})' in hp.text, f'header version must match latest CHANGES edition {CH["latest"]}'
    p = doc.add_paragraph(); p.add_run().add_break(WD_BREAK.PAGE)
    p = doc.add_paragraph(); runs(p, f'{CH["base"].split()[0]} 대비 변경 요약', size=15, color=NAVY, bold=True)
    para_border(p, 'bottom', 12, NAVY, 4)
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4)
    runs(p, f'기준판 {CH["base"]} → 최신판 {CH["latest"]}. 최신판 변경은 맨 위·굵게. 등급은 바뀐 수치의 근거 등급(A~D, 표지 뒤 안내 참조).',
         size=8.5, color=MUTED)
    t = doc.add_table(rows=1 + len(CH['rows']), cols=4)
    fix_table(t, fit([1.25, 1.0, 6.2, 1.45]))
    borders(t._tbl.tblPr, top=(12, NAVY), bottom=(12, NAVY), insideH=(4, RULE), left=None, right=None, insideV=None)
    for i, h in enumerate(('판', '장', '바뀐 점', '근거 등급')):
        shade(t.rows[0].cells[i], NAVY)
        cell_text(t.rows[0].cells[i], h, size=8, bold=True, color='FFFFFF', align=WD_ALIGN_PARAGRAPH.CENTER)
    for r, row in enumerate(CH['rows'], start=1):
        cells = t.rows[r].cells
        if row['latest']:
            for c in cells:
                shade(c, PANEL)
        vals = (f'{row["ver"]}\n{row["date"]}', row['chapter'], row['what'], row['grade'])
        for i, v in enumerate(vals):
            cell_text(cells[i], v, size=8, bold=row['latest'] or i == 0, color=NAVY if i == 0 else INK,
                      align=WD_ALIGN_PARAGRAPH.LEFT if i == 2 else WD_ALIGN_PARAGRAPH.CENTER)
    for line in CH['common']:
        source('· ' + line)

if CH['questions']:
    p = doc.add_paragraph(); runs(p, '검토 중 의문사항', size=15, color=NAVY, bold=True); p.paragraph_format.page_break_before = True  # no empty break paragraph: a full change page must not spill a blank page
    para_border(p, 'bottom', 12, NAVY, 4)
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4)
    runs(p, f'논리 구조를 다지기 위해 판마다 남기는 열린 질문({CH["questions_src"]} 등). 답이 본문에 반영되면 상태를 갱신하고 해결된 행은 뺌. 팀 공유용.',
         size=8.5, color=MUTED)
    t = doc.add_table(rows=1 + len(CH['questions']), cols=3)
    fix_table(t, fit([0.9, 5.2, 3.8]))
    borders(t._tbl.tblPr, top=(12, NAVY), bottom=(12, NAVY), insideH=(4, RULE), left=None, right=None, insideV=None)
    for i, h in enumerate(('장', '의문', '현재 상태 · 다음 조치')):
        shade(t.rows[0].cells[i], NAVY)
        cell_text(t.rows[0].cells[i], h, size=8, bold=True, color='FFFFFF', align=WD_ALIGN_PARAGRAPH.CENTER)
    for r, q in enumerate(CH['questions'], start=1):
        vals = (q['chapter'], q['question'], q['status'])
        for i, v in enumerate(vals):
            cell_text(t.rows[r].cells[i], v, size=8, bold=i == 0, color=NAVY if i == 0 else INK,
                      align=WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.LEFT)

# ---------- table of contents ----------
# page_break_before instead of an empty break paragraph: a full change-summary page must not spill a blank page
p = doc.add_paragraph(style='Heading 1'); runs(p, '목차'); para_border(p, 'bottom', 12, NAVY, 4)
p.paragraph_format.keep_with_next = True
p.paragraph_format.page_break_before = True
toc = doc.add_paragraph()
r = toc.add_run(); r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>'))
r = toc.add_run(); r._r.append(parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>'))
r = toc.add_run(); r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>'))
toc.add_run('목차는 Word에서 필드 업데이트 시 생성됩니다.')
r = toc.add_run(); r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>'))
# the TOC heading itself must not be listed: give it a plain style look but outline level body
p.style = doc.styles['Normal']; p.runs[0].font.size = Pt(15); p.runs[0].bold = True; p.runs[0].font.color.rgb = rgb(NAVY)

# ---------- content ----------
src = CONTENT.read_text(encoding='utf-8')
body = src[src.index('# ================= 1. Investment Highlights'):src.index('doc.save(OUT)')]
exec(compile(body, 'content', 'exec'), {**globals(), 'CHARTS': CHARTS})

for style in ('TOC 1', 'TOC 2', 'TOC 3'):
    try:
        st = doc.styles[style]
        style_font(st, {'TOC 1': 10, 'TOC 2': 9.5, 'TOC 3': 9}[style], style == 'TOC 1', {'TOC 1': NAVY, 'TOC 2': INK, 'TOC 3': MUTED}[style])
        st.paragraph_format.space_after = Pt(4)
    except KeyError:
        pass

doc.save(OUT)
print('saved', OUT, COUNT)
