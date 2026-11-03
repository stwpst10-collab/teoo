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
    '시장 분석': '디바이스 산업의 규제 경계는 사용 목적·사양·표현에 따라 갈리고, 대형 홈헬스케어는 정체기. 국내 성인 여성 41.6%가 심한 월경통이 있다고 답했고, 심한 증상 응답자 중 산부인과·비뇨의학과 진료 경험은 28.5%. 인접 시장 규모는 배경이며 합산하지 않음.',
    '사업 분석 및 투자포인트': '발목 착용형은 첫 검증 가설. 구매 목적을 셋으로 나눠 확인하고, 같은 고객에게서 복부·허리 수요가 확인되면 순서를 바꾸며, 문제·지불 이유가 약하면 시제품 이전에도 멈춤. "그날의"는 검증 위에 쌓을 모브랜드 비전.',
    '경쟁 분석': '생리통 착용형(Beurer·Myoovi·Livia), 복부 온열 착용(Stoov), 공통 본체·스트랩 구조(Hyperice·Therabody)가 이미 있어 형태·공용화는 해자가 아님. 에이피알은 전용 소모품이 아니라 기기와 화장품의 병용·교차판매 모델.',
    '재무계획': '기기 단독 손익분기는 약 6,286대로 3년차 계획 5,000대에서는 -0.45억 원. 패치는 교체·재구매가 확인되기 전 기본 시나리오에서 제외. 흑자 경로는 가격·원가·판매량·후속 제품 등 여러 가지. 모든 손익은 계획 시나리오.',
    '매출추정 및 사업가치': '심한 월경통 응답 여성 470만 명(추정)의 1%가 누적으로 본체를 사면 약 56억 원 규모(연간 아님, 예시). 실패 원인별 전환 경로를 정하고, 다음 판단에 실제로 쓰이는 기록만 자산으로 설명함.',
    '리스크 및 대응': '가장 큰 리스크는 차별화(이미 있는 착용형 제품 대비 선택 이유), 효능 표현·규제, 수요 미검증, IP 소유권.',
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
    p = doc.add_paragraph(style='Heading 3')
    runs(p, '■ ', color=ACCENT); runs(p, t)


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
    s1('핵심 지표')
    kv_table([('심한 월경통 응답', '성인 여성 41.6% (A)'), ('산부인과 진료 경험', '심한 증상 응답자 중 28.5% (A)'),
              ('약국 진통제 대처', '59.2% (A)'), ('심한 월경통 응답 여성(추정)', '약 470만~500만 명 (D)'),
              ('국내 홈뷰티 디바이스', '1.6조 원(2022) → 3.4조 원(2030E) (B)'), ('국내 안마 기기', '1.6조~2조 원, 정체기 (B)'),
              ('W케어 상품 수', '약 1만 8천 개, 2022년 대비 약 5배 (B)'), ('1% 침투 예시', '누적 본체 매출 약 56억 원 (D)')])


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
runs(hp, '그날의 류  |  산업·기업 분석 보고서  |  2026.11 (v8)', size=8, color=MUTED)
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
runs(p, '손을 멈추지 않는 케어 — 생리 전후 발목에서 시작하는 착용형 케어 브랜드', size=13, color=INK)
para_border(p, 'bottom', 18, NAVY, 10)
kv_table([('작성일', '2026년 11월 3일 (v8)'), ('형식', '산업분석 + 기업분석 풀리포트')])
callout('그날의 류는 "생리 지압기"가 아니라, 여성이 반복되는 일상 불편 때문에 하던 일을 멈추지 않도록 착용형 케어 제품을 탐색하는 '
        '브랜드다. 첫 제품은 생리 전후 발목 케어이며, 수동형과 전자형은 고객의 사용 상황·기능의 추가 가치·가격 수용성·사업성을 비교해 선택한다. '
        '결과에 따라 형태를 고치거나 바꾸거나 멈추며, "그날의" 시리즈 확장은 검증된 첫 제품 위에서만 추진한다. 다음 판단에 실제로 쓰일 고객·시험 기록과 그 사용 권리를 함께 남기는 것을 목표로 한다.',
        label='한 문장 논지')
for _ in range(4):
    doc.add_paragraph()
p = doc.add_paragraph(); runs(p, 'Compliance Notice', size=8, color=NAVY, bold=True)
p = doc.add_paragraph()
runs(p, '본 자료는 사내벤처 심사를 위한 내부 검토 자료로, 팀 지식 저장소(teoo)의 근거에 기반해 작성되었습니다. 수치 뒤의 (A)~(D)는 근거 등급이며 '
        'A 공시·정부 통계·규제 문서·논문 원문, B 리포트·언론 원문 인용, C 업체 주장·판매용 요약·AI 답변, D 팀 가설·내부 추정을 뜻합니다. '
        '추정·시나리오는 실적이나 검증 결과가 아니며, 제품의 의학적 효능을 주장하지 않습니다.', size=7.5, color=MUTED)

# ---------- table of contents ----------
p = doc.add_paragraph(); p.add_run().add_break(WD_BREAK.PAGE)
p = doc.add_paragraph(style='Heading 1'); runs(p, '목차'); para_border(p, 'bottom', 12, NAVY, 4)
p.paragraph_format.keep_with_next = True
toc = doc.add_paragraph()
r = toc.add_run(); r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>'))
r = toc.add_run(); r._r.append(parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText>'))
r = toc.add_run(); r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>'))
toc.add_run('목차는 Word에서 필드 업데이트 시 생성됩니다.')
r = toc.add_run(); r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>'))
# the TOC heading itself must not be listed: give it a plain style look but outline level body
p.style = doc.styles['Normal']; p.runs[0].font.size = Pt(15); p.runs[0].bold = True; p.runs[0].font.color.rgb = rgb(NAVY)

# ---------- content ----------
src = CONTENT.read_text(encoding='utf-8')
body = src[src.index('# ================= 1. Investment Highlights'):src.index('doc.save(OUT)')]
exec(compile(body, 'content', 'exec'), {**globals(), 'CHARTS': CHARTS})

for style in ('TOC 1', 'TOC 2'):
    try:
        st = doc.styles[style]
        style_font(st, 10 if style == 'TOC 1' else 9.5, style == 'TOC 1', NAVY if style == 'TOC 1' else INK)
        st.paragraph_format.space_after = Pt(4)
    except KeyError:
        pass

doc.save(OUT)
print('saved', OUT, COUNT)
