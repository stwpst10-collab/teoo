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
    '시장 분석': '19~49세 심한 PMS 응답 여성은 약 429만~458만명(D)이나 PMS 진료 인원은 2025년 1만 7,255명으로 10~49세 여성의 0.14%에 불과해(A), 수요 대부분이 의료 체계 밖에 있는 것으로 판단된다. 핵심 타깃 25~39세는 약 228만~239만명, 비약물 자가 대처 경험자(SAM)는 약 73만~188만명으로 추정되며(D), 첫해 구매자(SOM)는 별도 상향식으로 140~4,440명(기본 960명)을 목표로 한다. 생리 중(통증) 시장은 3만원대 전동기기(쿠팡 상위 67개 중앙값 34,900원, A)로 경쟁이 과밀한 반면, 생리 전 지압 결합 상품은 진열 공백 상태다. 확장은 같은 기기 심화(Oura 매출 +123%, A)를 우선하고, 기기 → 기기 확장(풀리오 −29%, 세라젬 −27%, A)은 단계 조건 충족 시에만 추진한다.',
    '사업 분석 및 투자포인트': '첫 판매 패키지는 발목 혈자리 지압 기기와 디바이스 전용 1회용 온열 패치(DEC-002, 조건부 채택)로, 생리 전 7일 착용 루틴을 겨냥한다. 생리 지연은 사업 착안 배경으로만 다루며, 효능은 표방하지 않는다(삼음교 지압 근거는 월경통 VAS −1점 수준, PMS는 RCT 5건, A). 계획가 13.2만원은 시장 기준점 3만원대와 격차가 커 가격 3안(2.99만·4.9만·9.9만원) 판매 페이지로 검증한다. 투자포인트는 생리 전 증상 수요, 손을 쓰지 않는 루틴, 기기+패치 구독, 진열 공백, 주기 축 확장의 5가지로 정리된다.',
    '경쟁 분석': '생리통 착용형(Beurer·Myoovi·Livia·Ova+), 복부 온열 착용(Stoov), 공통 본체·스트랩 구조(Hyperice·Therabody)가 이미 존재해 형태·공용화는 경쟁우위로 주장하지 않는다. 생리통 검색 상위 기기 가격은 3만원대(A)이며 국내 7사 제품은 상위 91개에 포함되지 않는다. 7사는 매출의 15~38%를 광고에 쓰는 온라인 D2C 구조다. 기기+소모품 모델은 DEC-002로 조건부 채택했으나 소모품 잠금은 고객 반발 위험이 있어 "기기만으로도 충분해요. 패치는 선택" 원칙을 병행한다.',
    '재무계획': '수익 모델은 기기 1회 판매 + 1회용 온열 패치 구독(DEC-002)이다. 3년차 세전 손익은 계획안(패치 포함) 0.75억원, 보수안(기기 단독) −0.45억원이며, 기기 단독 손익분기는 약 6,286대다(린캔버스 팀 가정, D. CAC 1.5만원·환불 0% 전제). 패치 재구매율이 미측정인 만큼 투자 판단은 보수안 기준이 적절하다고 판단된다. 모든 손익은 계획 시나리오이며 재무 시뮬레이션은 반영 전이다.',
    '매출추정 및 사업가치': '첫해 판매 규모는 시장 깔때기 기준 구매자 140~4,440명(기본 960명), 매출 0.12억~5.73억원(소비자가, D)으로 추정된다. 방문 수는 미측정으로, 100일 검증에서 실측해 범위를 좁힐 계획이다. 규모 예시로 심한 PMS 응답 여성 약 450만명의 1%가 누적 구매 시 본체 매출은 약 54억원이다(1%는 설명용 가정, D). 외부 근거 감사 대안(침투 분모 SAM, 패치 재구매 20% 등)은 재무 반영 시 확정한다.',
    '리스크 및 대응': '주요 리스크는 효능 근거 제한과 표현 규제, 낮은 사용 빈도(생리 전 루틴 7일, D), 저가 경쟁, 채널 의존, 상표 충돌이다. 안전 측면에서는 임신 중 사용 금지와 생리가 늦은 고객의 임신 확인 안내, 온열 패치의 피부 42°C 이하 조건(DEC-002 GO 조건)을 설계 기준으로 둔다. 확인 시점은 날짜가 아닌 단계(100일 검증·확정 견적·식약처 질의·변리사 자문)로 관리한다.',
    '결론': '그날의 류는 PMS 일주일이라는 미개척 영역에서 호르몬제 없이 관리하는 첫 루틴 브랜드를 지향한다. 검증할 핵심 가설은 생리 전 루틴 수요와 가격, 패치 재구매, 착용 제품의 선택 이유 세 가지이며, 각각 판매 페이지 전환율·90일 재구매율·2주기 연속 사용률로 확인한다. 확장은 고객의 주기를 축으로 단계 조건을 충족할 때만 진행한다.',
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
    kv_table([('심한 PMS 경험률 (19~39세)', '48.3% (2025, 2022년 51.1%, 질병관리청, A)'), ('PMS 진료 인원 (N94.3)', '1만 7,255명 (2025, 10~49세 여성의 0.14%, 심평원, A)'),
              ('TAM: 심한 PMS 응답 여성', '약 429만~458만명 (19~49세, 연령별 비율 × 인구, D)'), ('핵심 타깃 (25~39세)', '약 228만~239만명 (D)'),
              ('SAM: 비약물 자가 대처 경험자', '약 73만~188만명 (TAM × 17~41%, 대리지표, D)'), ('SOM: 첫해 구매자', '140~4,440명, 기본 960명 (0.12억~5.73억원, D)'),
              ('약국 진통제 대처', '59.2% (심한 월경 증상 성인 1,266명 중, A)'), ('생리통 검색 기기 가격 기준점', '34,900원 (쿠팡 상위 전동기기 67개 중앙값, A)'),
              ('소형 안마기 공시 4사 매출', '1,930억원 (2025, A 합산, 비공시 제외 하한)'), ('W케어 상품 수', '약 1만 8천개, 2022년 대비 약 5배 (B)'),
              ('수익 모델', '기기 1회 판매 + 1회용 온열 패치 구독 (DEC-002 조건부 채택)'), ('참고: 심한 월경통 응답 여성', '약 406만명 (성인 41.6% × 19~49세 인구, D)')])


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
runs(hp, '그날의 류  |  산업·기업 분석 보고서  |  2026.11 (v10.12)', size=8, color=MUTED)
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
runs(p, 'PMS 일주일, 아무도 맡지 않은 시장 — 두 손을 자유롭게! 발목 혈자리 지압기 & 전용 패치 형태의 생리전 케어 디바이스', size=13, color=INK)
para_border(p, 'bottom', 18, NAVY, 10)
kv_table([('작성일', '2026년 11월 7일 (v10.12)'), ('형식', '산업분석 + 기업분석 풀리포트')])
callout('그날의 류는 생리 전 7일(PMS)을 겨냥한 착용형 셀프케어 디바이스 브랜드다. 19~39세 여성의 48.3%가 심한 PMS를 경험하나 진료율은 0.14%에 불과해, '
        '약 외에 일상에서 쓸 관리 수단이 사실상 부재한 것으로 판단된다. 첫 제품은 혈자리를 몰라도, 손을 쓰지 않아도 발목에 착용하면 되는 혈자리 지압 기기와 디바이스 전용 패치이며, '
        '기기 1회 판매 + 패치 구독(DEC-002)으로 반복 매출을 확보할 계획이다. 확장은 부위가 아닌 고객의 주기를 축으로, 단계별 진입 조건을 충족할 때만 진행한다. 효능은 표방하지 않는다.',
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
