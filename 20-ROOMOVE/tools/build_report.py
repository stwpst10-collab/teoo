"""ROOMOVE 보고서 빌더 (그날의 류 해커톤·최종 보고서 양식을 그대로 씀).

    python 20-ROOMOVE/tools/build_report.py <content.py> <charts_dir> <out.docx>

content.py 위쪽의 META·SUMMARY·OVERVIEW 값과, "# ================= 1. Investment Highlights" 표시부터
"doc.save(OUT)" 앞까지의 본문 블록을 읽는다. 서식은 venture/00-운영/보고서 워드 작업 기준.md 6장 값.
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
TEXT_W_CM = 18.6  # A4 21cm - 좌우 여백 1.2cm x 2
TWIP = 567  # per cm

doc = docx.Document()
sec = doc.sections[0]
# 작업 기준 6장(참고 양식 「바이오뱅크힐링.docx」): A4, 좌우 1.2cm, 위 2.2cm, 아래 1.5cm
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(1.2)
sec.top_margin, sec.bottom_margin = Cm(2.2), Cm(1.5)
sec.header_distance, sec.footer_distance = Cm(1.0), Cm(0.7)
sec.different_first_page_header_footer = True

LINE = 1.075          # 본문·제목·글머리 줄간격 배수
BODY_IND = 0.5        # 본문 왼쪽 들여쓰기(cm), 표도 같은 위치에서 시작
TABLE_W_CM = TEXT_W_CM - BODY_IND


def rgb(h):
    return RGBColor.from_string(h)


def style_font(st, size, bold=False, color=INK):
    st.font.name = FONT
    st.element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'), FONT)
    st.font.size = Pt(size); st.font.bold = bold; st.font.color.rgb = rgb(color)


def no_keep(pf):
    # 작업 기준 6장: '다음 문단과 함께'·'현재 문단을 나누지 않음'·'앞에서 쪽 나누기' 금지(편집 기호의 검은 네모)
    pf.keep_with_next = False; pf.keep_together = False; pf.page_break_before = False


normal = doc.styles['Normal']
style_font(normal, 9)
npf = normal.paragraph_format
npf.line_spacing, npf.space_before, npf.space_after = LINE, Pt(0), Pt(4)
npf.left_indent, npf.alignment = Cm(BODY_IND), WD_ALIGN_PARAGRAPH.JUSTIFY
no_keep(npf)
# 작업 기준 6장(2026-11-25): 한글 글자 단위 줄 바꿈(단어 잘림 허용) — 양쪽 맞춤 간격 벌어짐 방지
normal.element.get_or_add_pPr().append(parse_xml(f'<w:wordWrap {nsdecls("w")} w:val="0"/>'))
for name, size, before, after in (('Heading 1', 10, 12, 4), ('Heading 2', 9.5, 6, 2)):
    st = doc.styles[name]
    style_font(st, size, True, INK)
    pf = st.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(before), Pt(after), LINE
    pf.left_indent, pf.alignment = Cm(0), WD_ALIGN_PARAGRAPH.LEFT
    no_keep(pf)


# ---------- 글머리 스타일(1단계 ▪ Wingdings, 2단계 -) : 워드 번호 매기기로 만든다 ----------
def _bullet_styles():
    numbering = doc.part.numbering_part.element
    specs = [('글머리 1단계', 'Bullet1', chr(0xF0A7), 'Wingdings', 0.5, 1.0, 4, 4, 901),
             ('글머리 2단계', 'Bullet2', '-', FONT, 1.1, 1.6, 0, 0, 902)]
    for name, sid, char, font, num_cm, text_cm, before, after, nid in specs:
        left, hang = int(text_cm * TWIP), int((text_cm - num_cm) * TWIP)
        numbering.insert(0, parse_xml(
            f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{nid}"><w:multiLevelType w:val="singleLevel"/>'
            f'<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:pStyle w:val="{sid}"/>'
            f'<w:lvlText w:val="{char}"/><w:lvlJc w:val="left"/><w:suff w:val="tab"/>'
            f'<w:pPr><w:tabs><w:tab w:val="num" w:pos="{left}"/></w:tabs><w:ind w:left="{left}" w:hanging="{hang}"/></w:pPr>'
            f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:hint="default"/><w:sz w:val="18"/></w:rPr>'
            f'</w:lvl></w:abstractNum>'))
        numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="{nid}"><w:abstractNumId w:val="{nid}"/></w:num>'))
        st = doc.styles.add_style(name, 1)
        st.element.set(qn('w:styleId'), sid)
        st.base_style = normal
        style_font(st, 9)
        pf = st.paragraph_format
        pf.space_before, pf.space_after, pf.line_spacing = Pt(before), Pt(after), LINE
        pf.left_indent, pf.first_line_indent = Cm(text_cm), Cm(-(text_cm - num_cm))
        no_keep(pf)
        st.element.get_or_add_pPr().append(parse_xml(
            f'<w:numPr {nsdecls("w")}><w:ilvl w:val="0"/><w:numId w:val="{nid}"/></w:numPr>'))


_bullet_styles()

# ---------- 쪽 아래 각주: 원고의 [^내용]을 워드 각주로 바꾼다 (작업 기준 4장) ----------
from docx.opc.constants import RELATIONSHIP_TYPE as RT  # noqa: E402
from docx.opc.packuri import PackURI  # noqa: E402
from docx.opc.part import XmlPart  # noqa: E402

doc.styles.element.append(parse_xml(
    f'<w:style {nsdecls("w")} w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/>'
    f'<w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>'
    f'<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="0" w:firstLine="0"/><w:jc w:val="both"/></w:pPr>'
    f'<w:rPr><w:color w:val="{INK}"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>'))
doc.styles.element.append(parse_xml(
    f'<w:style {nsdecls("w")} w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/>'
    f'<w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'))
_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
_SEP = '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="0"/></w:pPr><w:r><w:{0}/></w:r></w:p>'
FOOT = XmlPart(PackURI('/word/footnotes.xml'),
               'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
               parse_xml(f'<w:footnotes {nsdecls("w")}>'
                         f'<w:footnote w:type="separator" w:id="-1">{_SEP.format("separator")}</w:footnote>'
                         f'<w:footnote w:type="continuationSeparator" w:id="0">{_SEP.format("continuationSeparator")}</w:footnote>'
                         f'</w:footnotes>'), doc.part.package)
doc.part.relate_to(FOOT, RT.FOOTNOTES)
doc.settings.element.append(parse_xml(f'<w:footnotePr {nsdecls("w")}><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'))
FN = {'n': 0}
FN_RE = re.compile(r'\[\^([^\]]+)\]')


def footnote(p, text):
    FN['n'] += 1
    n = FN['n']
    fn = parse_xml(f'<w:footnote {nsdecls("w")} w:id="{n}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
                   f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
                   f'<w:r><w:t xml:space="preserve"> </w:t></w:r></w:p></w:footnote>')
    fn.findall(f'.//{{{_W}}}t')[-1].text = ' ' + text.strip()
    FOOT.element.append(fn)
    p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="{n}"/></w:r>'))


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
    for k, chunk in enumerate(FN_RE.split(text)):
        if k % 2:  # [^...] → 쪽 아래 각주
            footnote(p, chunk)
            continue
        for part in re.split(r'(\*\*[^*]+\*\*)', chunk):
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


def fit(widths, total_cm=None):
    tw = int((total_cm or TABLE_W_CM) * TWIP)
    ws = [int(tw * w / sum(widths)) for w in widths]
    ws[-1] += tw - sum(ws)
    return ws


def fix_table(t, ws, indent_cm=BODY_IND):
    tblpr = t._tbl.tblPr
    for tag in ('w:tblW', 'w:tblLayout', 'w:tblInd', 'w:jc'):
        for x in tblpr.findall(qn(tag)):
            tblpr.remove(x)
    tblpr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{sum(ws)}" w:type="dxa"/>'))
    tblpr.append(parse_xml(f'<w:tblInd {nsdecls("w")} w:w="{int(indent_cm * TWIP)}" w:type="dxa"/>'))
    tblpr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    tblpr.append(parse_xml(f'<w:tblCellMar {nsdecls("w")}><w:top w:w="28" w:type="dxa"/><w:left w:w="85" w:type="dxa"/>'
                           f'<w:bottom w:w="28" w:type="dxa"/><w:right w:w="85" w:type="dxa"/></w:tblCellMar>'))
    grid = t._tbl.tblGrid
    for i, g in enumerate(grid.findall(qn('w:gridCol'))):
        g.set(qn('w:w'), str(ws[i]))
    for row in t.rows:
        for i, c in enumerate(row.cells):
            c.width = ws[i]
            tcw = c._tc.get_or_add_tcPr().find(qn('w:tcW'))
            tcw.set(qn('w:type'), 'dxa'); tcw.set(qn('w:w'), str(ws[i]))


def cell_text(c, text, size=8, bold=False, color=INK, align=WD_ALIGN_PARAGRAPH.LEFT):
    c.text = ''
    lines = str(text).split('\n')
    for k, ln in enumerate(lines):
        p = c.paragraphs[0] if k == 0 else c.add_paragraph()
        p.alignment = align
        pf = p.paragraph_format
        pf.space_after, pf.space_before, pf.line_spacing, pf.left_indent = Pt(1), Pt(1), 1.0, Cm(0)
        runs(p, ln, size=size, color=color, bold=bold)
    c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


COUNT = {'표': 0, '그림': 0}


def caption(kind, title):
    COUNT[kind] += 1
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(6), Pt(3)
    p.paragraph_format.widow_control = True  # 그림 제목+그림 두 줄이 쪽 사이에서 갈라지지 않게(기본값 명시)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT  # 양쪽 맞춤이면 줄바꿈 앞 제목 글자가 벌어진다
    runs(p, f'[{kind} {COUNT[kind]}] ', size=9, color=INK, bold=True)
    runs(p, title, size=9, color=INK, bold=True)


def source(text):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(2), Pt(8), 1.0
    pf.left_indent, pf.alignment = Cm(BODY_IND), WD_ALIGN_PARAGRAPH.LEFT
    runs(p, text, size=7.5, color=MUTED)


def page_break():
    """장 앞 쪽 나누기: 앞 장 마지막 문장 끝에 붙인다. 앞이 표·그림·빈 줄이면 1pt 높이 줄을 따로 둔다(작업 기준 6장)."""
    body = doc.element.body
    last = body[-1] if body[-1].tag != qn('w:sectPr') else body[-2]
    if last.tag == qn('w:p') and ''.join(last.itertext()).strip() and last.find(f'.//{{{_W}}}drawing') is None:
        last.append(parse_xml(f'<w:r {nsdecls("w")}><w:br w:type="page"/></w:r>'))
        # 양쪽 맞춤은 쪽 나누기로 끝나는 마지막 줄까지 늘려 글자 사이가 벌어짐 → 이 문단만 왼쪽 맞춤
        ppr = last.get_or_add_pPr()
        for jc in ppr.findall(qn('w:jc')):
            ppr.remove(jc)
        ppr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="left"/>'))
        return
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(0), Pt(0), Pt(1)
    r = p.add_run(); r.font.size = Pt(1); r.add_break(WD_BREAK.PAGE)


# ---------- content primitives (names used by the content script) ----------
H = {'h1': 0, 'h2': 0}
_META_SRC = CONTENT.read_text(encoding='utf-8')
_ns = {}
exec(compile(_META_SRC[:_META_SRC.index('# ================= 1. Investment Highlights')], 'meta', 'exec'), _ns)
META, SUMMARY, OVERVIEW = _ns['META'], _ns['SUMMARY'], _ns['OVERVIEW']


def callout(text, label='핵심 요약'):
    t = doc.add_table(rows=1, cols=1)
    fix_table(t, fit([1]))
    c = t.cell(0, 0); shade(c, PANEL)
    borders(t._tbl.tblPr, top=None, bottom=None, left=(24, ACCENT), right=None, insideH=None, insideV=None)
    c.text = ''
    p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(2); p.paragraph_format.left_indent = Cm(0)
    runs(p, label, size=8.5, color=NAVY, bold=True)
    p2 = c.add_paragraph(); pf = p2.paragraph_format
    pf.space_after, pf.line_spacing, pf.left_indent = Pt(1), LINE, Cm(0)
    runs(p2, text, size=9, color=INK)
    sp = doc.add_paragraph(); sp.paragraph_format.space_after = Pt(0); sp.paragraph_format.line_spacing = Pt(6)


def h1(t):
    H['h1'] += 1; H['h2'] = 0
    page_break()  # 앞 장(또는 목차) 끝에 쪽 나누기 — '앞에서 쪽 나누기' 설정은 쓰지 않는다
    p = doc.add_paragraph(style='Heading 1')
    runs(p, f'{H["h1"]}. {t}')
    para_border(p, 'bottom', 8, INK, 3)
    if t == 'Investment Highlights':
        overview()
    if t in SUMMARY:
        callout(SUMMARY[t])


def h2(t):
    H['h2'] += 1
    p = doc.add_paragraph(style='Heading 2')
    # 작업 기준 1장: 장 아래 절이 하나뿐이면 절 번호를 붙이지 않는다
    one = H2N.get(H['h1'], 0) <= 1
    runs(p, t if one else f'{H["h1"]}.{H["h2"]} {t}')


def s1(t):
    # 절 안의 굵은 소제목 줄(파란 라벨·제목 스타일 아님, 목차에 넣지 않음)
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(4), Pt(2)
    runs(p, t, bold=True)


def bullet(text, level):
    p = doc.add_paragraph(style='글머리 1단계' if level == 0 else '글머리 2단계')
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
    # 표 제목은 표의 첫 머리글 행에 넣는다: '다음 문단과 함께' 없이도 제목만 쪽 끝에 남지 않고, 표가 넘어가면 제목·머리말이 반복된다
    COUNT['표'] += 1
    n = len(header)
    ws = fit(widths or [1] * n)
    prose = [any(len(str(r[i])) > 16 for r in rows) for i in range(n)]
    t = doc.add_table(rows=2 + len(rows), cols=n)
    fix_table(t, ws)
    borders(t._tbl.tblPr, top=None, bottom=(12, NAVY), insideH=(4, RULE), left=None, right=None, insideV=None)
    cap_row = t.rows[0]
    cap_row._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    cap_row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
    cc = cap_row.cells[0].merge(cap_row.cells[-1]) if n > 1 else cap_row.cells[0]
    cell_text(cc, f'[표 {COUNT["표"]}] ' + re.sub(r'^\[도표 \d+\] ', '', title), size=9, bold=True)
    cc.paragraphs[0].paragraph_format.space_before = Pt(6)
    cc.paragraphs[0].paragraph_format.space_after = Pt(3)
    cc._tc.get_or_add_tcPr().append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="nil"/><w:left w:val="nil"/>'
                                              f'<w:right w:val="nil"/><w:bottom w:val="single" w:sz="12" w:space="0" w:color="{NAVY}"/></w:tcBorders>'))
    cc._tc.get_or_add_tcPr().append(parse_xml(f'<w:tcMar {nsdecls("w")}><w:left w:w="0" w:type="dxa"/></w:tcMar>'))
    hdr = t.rows[1]
    hdr._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    for i, h in enumerate(header):
        shade(hdr.cells[i], NAVY)
        cell_text(hdr.cells[i], h, bold=True, color='FFFFFF', align=WD_ALIGN_PARAGRAPH.CENTER)
    for r, row in enumerate(rows, start=1):
        tr = t.rows[r + 1]
        tr._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))  # 행이 쪽 사이에서 갈라지지 않게
        for i in range(n):
            v = row[i] if i < len(row) else ''
            c = tr.cells[i]
            if r % 2 == 0:
                shade(c, ZEBRA)
            align = WD_ALIGN_PARAGRAPH.LEFT if (i == 0 or prose[i]) else WD_ALIGN_PARAGRAPH.CENTER
            if i == 0 and not prose[0]:
                align = WD_ALIGN_PARAGRAPH.CENTER
            cell_text(c, v, bold=(i == 0), color=NAVY if i == 0 else INK, align=align)
    source(src)


def figure(title, png, src, width_cm=15.0):
    # 그림 제목과 그림을 한 문단(줄바꿈)으로 묶는다: '다음 문단과 함께' 없이도 두 줄짜리 문단은
    # 워드의 줄 분리 방지(기본값)로 같은 쪽에 놓여 제목만 쪽 끝에 남지 않는다
    caption('그림', re.sub(r'^\[도표 \d+\] ', '', title))
    p = doc.paragraphs[-1]
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_break()
    p.add_run().add_picture(str(png), width=Cm(min(width_cm, TABLE_W_CM)))
    source(src)


def kv_table(pairs, label_w=3.3):
    """Two pairs per row: label | value | label | value."""
    rows = [pairs[i:i + 2] for i in range(0, len(pairs), 2)]
    t = doc.add_table(rows=len(rows), cols=4)
    vw = (TABLE_W_CM - 2 * label_w) / 2
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


def overview():
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(3)
    runs(p, '핵심 지표', bold=True)  # 제목 스타일 아님: 목차에 넣지 않는다
    kv_table(OVERVIEW['pairs'])
    source(OVERVIEW['source'])


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
runs(hp, META['header'], size=8, color=MUTED)
para_border(hp, 'bottom', 4, RULE, 3)
fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
page_field(fp)

# ---------- cover (해커톤 보고서 표지 형식) ----------
p = doc.add_paragraph(); p.paragraph_format.left_indent = Cm(0)
runs(p, META['kicker'], size=10, color=ACCENT, bold=True)
p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(6); p.paragraph_format.line_spacing = 1.0
p.paragraph_format.left_indent = Cm(0)
runs(p, META['title'], size=30, color=INK, bold=True)
para_border(p, 'bottom', 12, ACCENT, 8)
for line in META['lead']:
    p = doc.add_paragraph(); runs(p, line)
t = doc.add_table(rows=1 + len(META['cover_rows']), cols=2)
fix_table(t, fit([1, 1.6]))
borders(t._tbl.tblPr, top=(12, NAVY), bottom=(12, NAVY), insideH=(4, RULE), left=None, right=None, insideV=None)
for i, h in enumerate(META['cover_head']):
    shade(t.rows[0].cells[i], PANEL)
    cell_text(t.rows[0].cells[i], h, bold=True, color=NAVY)
for r, (k, v) in enumerate(META['cover_rows'], start=1):
    cell_text(t.rows[r].cells[0], k)
    cell_text(t.rows[r].cells[1], v)
p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(4); runs(p, META['notice'])
CH = {'questions': META.get('questions', []), 'questions_src': META.get('questions_src', '')}

if CH['questions']:
    page_break()
    p = doc.add_paragraph(); runs(p, '검토 중 의문사항', size=10, bold=True)
    para_border(p, 'bottom', 8, INK, 3); p.paragraph_format.left_indent = Cm(0)
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4)
    runs(p, f'논리 구조를 다지기 위해 판마다 남기는 열린 질문({CH["questions_src"]}). 답이 본문에 반영되면 상태를 갱신하고 해결된 행은 뺌.',
         size=8.5, color=MUTED)
    t = doc.add_table(rows=1 + len(CH['questions']), cols=3)
    fix_table(t, fit([0.9, 5.2, 3.8]))
    borders(t._tbl.tblPr, top=(12, NAVY), bottom=(12, NAVY), insideH=(4, RULE), left=None, right=None, insideV=None)
    for i, h in enumerate(('장', '의문', '현재 상태 · 다음 조치')):
        shade(t.rows[0].cells[i], NAVY)
        cell_text(t.rows[0].cells[i], h, size=8, bold=True, color='FFFFFF', align=WD_ALIGN_PARAGRAPH.CENTER)
    for r, q in enumerate(CH['questions'], start=1):
        for i, v in enumerate(q):
            cell_text(t.rows[r].cells[i], v, size=8, bold=i == 0, color=NAVY if i == 0 else INK,
                      align=WD_ALIGN_PARAGRAPH.CENTER if i == 0 else WD_ALIGN_PARAGRAPH.LEFT)

# ---------- table of contents ----------
page_break()
p = doc.add_paragraph(); runs(p, '목차', size=10, bold=True); para_border(p, 'bottom', 8, INK, 3)
p.paragraph_format.left_indent = Cm(0)
toc = doc.add_paragraph()
r = toc.add_run(); r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>'))
r = toc.add_run(); r._r.append(parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>'))
r = toc.add_run(); r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>'))
toc.add_run('목차는 Word에서 필드 업데이트 시 생성됩니다.')
r = toc.add_run(); r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>'))

# ---------- content ----------
src = CONTENT.read_text(encoding='utf-8')
body = src[src.index('# ================= 1. Investment Highlights'):src.index('doc.save(OUT)')]
PHOTOS = CHARTS
# 장마다 절(h2) 개수를 미리 센다: 절이 하나뿐인 장은 절 번호를 붙이지 않는다
H2N, _ch = {}, 0
for line in body.splitlines():
    if line.startswith("h1("):
        _ch += 1; H2N[_ch] = 0
    elif line.startswith("h2(") and _ch:
        H2N[_ch] += 1
exec(compile(body, 'content', 'exec'), {**globals(), 'CHARTS': CHARTS, 'PHOTOS': PHOTOS})

for style in ('TOC 1', 'TOC 2'):  # 목차 9pt, 1줄 간격 (작업 기준 6장)
    try:
        st = doc.styles[style]
    except KeyError:
        st = doc.styles.add_style(style, 1); st.base_style = normal
    style_font(st, 9, style == 'TOC 1', INK)
    pf = st.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(0), Pt(2), 1.0
    pf.left_indent = Cm(0 if style == 'TOC 1' else 0.4)
    no_keep(pf)

doc.save(OUT)
print('saved', OUT, COUNT, 'footnotes', FN['n'])
