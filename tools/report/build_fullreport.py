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
        return
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(0), Pt(0), Pt(1)
    r = p.add_run(); r.font.size = Pt(1); r.add_break(WD_BREAK.PAGE)


# ---------- content primitives (names used by the content script) ----------
H = {'h1': 0, 'h2': 0}
SUMMARY = {  # 장 첫머리 요약 상자. 출처는 본문 각주에 있으므로 여기에는 등급·출처 괄호를 쓰지 않는다
    '시장 분석': '19~49세 심한 PMS 응답 여성은 약 429만~458만명으로 추정되나, PMS 진료 인원은 2025년 1만 7,255명으로 10~49세 여성의 0.14%에 불과해 수요 대부분이 의료 체계 밖에 있다고 판단. 핵심 타깃 25~39세는 약 228만~239만명, 비약물 자가 대처 경험자(유효 시장)는 약 73만~188만명으로 추정되며, 첫해 구매자는 별도의 상향식 계산으로 140~4,440명(기본 960명)이 목표. 생리 중(통증) 시장은 3만원대 전동 기기(쿠팡 상위 67개 중앙값 34,900원)로 경쟁이 과밀한 반면, 생리 전 지압 결합 상품은 진열 공백 상태. 확장은 같은 기기 심화(Oura 매출 +123%)를 우선하고, 다른 기기로의 확장(풀리오 −29%, 세라젬 −27%)은 단계 조건을 충족할 때만 추진. 디바이스 확장의 근거는 시장 크기가 아닌 올리브영 고객·채널에 기기를 얹는 효과로, 측정 전 가설이며 채널은 협의 전.',
    '사업 분석 및 투자포인트': '첫 판매 패키지는 발목 혈자리 지압 기기와 디바이스 전용 1회용 온열 패치(조건부 채택)로, 생리 전 7일 착용 루틴이 목표. 생리 지연은 사업 착안 배경으로만 다루며 효능은 표방하지 않음. 삼음교 지압 근거는 생리통 점수 약 1점(10점 척도) 감소 수준이고, PMS 대상 임상은 5건에 불과. 계획가 13.2만원은 시장 기준점 3만원대와 격차가 커 가격 3안(2.99만·4.9만·9.9만원) 판매 페이지로 검증 예정. 투자포인트는 생리 전 증상 수요, 손을 쓰지 않는 루틴, 기기+패치 구독, 진열 공백, 주기 축 확장의 5가지.',
    '경쟁 분석': '생리통 착용형(Beurer·Myoovi·Livia·Ova+), 복부 온열 착용(Stoov), 공통 본체·스트랩 구조(Hyperice·Therabody)가 이미 존재해 형태·공용화는 경쟁우위로 보지 않음. 해외에는 PMS를 표방한 뇌·귀 전기 자극 기기(Nettle·Lutea·OhmBody, 약 80만~100만원)가 출시됐으나, 학술지에 실린 유일한 임상에서 가짜 기기와의 차이가 뚜렷하지 않았음. 생리통 검색 상위 기기 가격은 3만원대이며 국내 7사 제품은 상위 91개에 없음. 7사는 매출의 15~38%를 광고에 쓰는 온라인 자사몰 직접 판매 구조. 기기+소모품(면도기-면도날) 모델은 조건부 채택했으나 소모품 락인은 고객 반발 위험이 있어 "기기만으로도 충분해요. 패치는 선택" 원칙을 병행.',
    '재무계획': '수익 모델은 기기 1회 판매 + 1회용 온열 패치 구독. 3년차 세전 손익은 계획안(패치 포함) 0.75억원, 보수안(기기 단독) −0.45억원이며, 기기 단독 손익분기는 약 6,286대(린캔버스 팀 가정, 고객 획득 비용 1.5만원·환불 0% 전제). 패치 재구매율이 아직 측정 전인 만큼 투자 판단은 보수안 기준이 적절하다고 판단. 모든 손익은 계획 시나리오이며 재무 시뮬레이션은 반영 전.',
    '매출추정 및 사업가치': '첫해 판매 규모는 시장 깔때기 기준 구매자 140~4,440명(기본 960명), 매출 0.12억~5.73억원(소비자가, 당사 추정)으로 추정. 방문 수는 측정 전으로, 100일 검증에서 실측해 범위를 좁힐 계획. 규모 예시로 심한 PMS 응답 여성 약 450만명의 1%가 누적 구매하면 본체 매출은 약 54억원(1%는 설명용 가정). 외부 근거 점검에서 나온 대안(침투 분모를 유효 시장으로, 패치 재구매 20% 등)은 재무 반영 시 확정 예정.',
    '리스크 및 대응': '주요 리스크는 효능 근거 제한과 표현 규제, 낮은 사용 빈도(생리 전 루틴 7일, 팀 가정), 저가 경쟁, 채널 의존, 상표 충돌. 안전 측면에서는 임신 중 사용 금지와 생리가 늦은 고객의 임신 확인 안내, 온열 패치의 피부 42°C 이하 조건(패치 채택의 진행 조건)을 설계 기준으로 설정. 확인 시점은 날짜가 아닌 단계(100일 검증·확정 견적·식약처 질의·변리사 자문)로 관리.',
    '결론': '그날의 류는 PMS 일주일이라는 미개척 영역에서 호르몬제 없이 관리하는 첫 루틴 브랜드를 지향. 검증할 핵심 가설은 생리 전 루틴 수요와 가격, 패치 재구매, 착용 제품의 선택 이유 세 가지이며, 각각 판매 페이지 전환율·90일 재구매율·2주기 연속 사용률로 확인 예정. 확장은 고객의 주기를 축으로 단계 조건을 충족할 때만 진행.',
}


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
    kv_table([('심한 PMS 경험률 (19~39세)', '48.3% (2025, 2022년 51.1%)'), ('PMS 진료 인원', '1만 7,255명 (2025, 10~49세 여성의 0.14%)'),
              ('전체 시장: 심한 PMS 응답 여성', '약 429만~458만명 (19~49세, 당사 추정)'), ('핵심 타깃 (25~39세)', '약 228만~239만명 (당사 추정)'),
              ('유효 시장: 비약물 자가 대처 경험자', '약 73만~188만명 (전체 시장 × 17~41%, 당사 추정)'), ('첫해 구매자 목표', '140~4,440명, 기본 960명 (0.12억~5.73억원)'),
              ('약국 진통제 대처', '59.2% (심한 월경 증상 성인 1,266명 중)'), ('생리통 검색 기기 가격 기준점', '34,900원 (쿠팡 상위 전동 기기 67개 중앙값)'),
              ('소형 안마기 공시 4사 매출', '1,930억원 (2025, 비공시 업체 제외 하한)'), ('W케어 상품 수', '약 1만 8천개, 2022년 대비 약 5배'),
              ('수익 모델', '기기 1회 판매 + 1회용 온열 패치 구독 (조건부 채택)'), ('참고: 심한 월경통 응답 여성', '약 406만명 (성인 41.6% × 19~49세 인구, 당사 추정)')])
    source('자료: 질병관리청 「한국 여성의 생애주기별 성·생식건강조사」(2022·2025), 건강보험심사평가원 질병통계(N94.3, 2025), 행정안전부 주민등록인구(2026.10), '
           '쿠팡 판매 페이지 관측(2026.11.04), 각 사 감사보고서(2025), 바이라인네트워크(2025.03), 당사 추정')


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
runs(hp, '그날의 류  |  산업·기업 분석 보고서  |  2026.11 (v10.14)', size=8, color=MUTED)
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
kv_table([('작성일', '2026년 11월 8일 (v10.14)'), ('형식', '산업분석 + 기업분석 풀리포트')])
callout('그날의 류는 생리 전 7일(PMS)을 겨냥한 착용형 셀프케어 디바이스 브랜드. 19~39세 여성의 48.3%가 심한 PMS를 경험하나 진료율은 0.14%에 불과해, '
        '약 외에 일상에서 쓸 관리 수단이 사실상 없다고 판단. 첫 제품은 혈자리를 몰라도, 손을 쓰지 않아도 발목에 착용하면 되는 혈자리 지압 기기와 디바이스 전용 패치이며, '
        '기기 1회 판매 + 패치 구독으로 반복 매출을 확보할 계획. PMS는 첫 진입점이며, 확장은 부위가 아닌 고객의 주기를 축으로 단계별 진입 조건을 충족할 때만 진행. 효능은 표방하지 않음.',
        label='한 문장 논지')
for _ in range(4):
    doc.add_paragraph()
p = doc.add_paragraph(); runs(p, 'Compliance Notice', size=8, color=NAVY, bold=True)
p = doc.add_paragraph()
runs(p, '본 자료는 사내벤처 심사를 위한 내부 검토 자료로, 팀 지식 저장소(teoo)의 근거에 기반해 작성되었습니다. 출처는 각 쪽 아래 각주와 표·그림 아래 「자료」에 적었고, '
        '근거의 성격은 공식 자료(공시·정부 통계·규제 문서), 학술지 논문, 언론 보도, 업체 자료, 직접 관측, 당사 추정으로 구분해 밝혔습니다. '
        '추정·시나리오는 실적이나 검증 결과가 아니며, 제품의 의학적 효능을 주장하지 않습니다.', size=7.5, color=MUTED)

# ---------- v8 대비 변경 요약 (one page after the cover; data at the top of report_content.py) ----------
sys.path.insert(0, str(Path(__file__).resolve().parent))
from changes import load_changes  # noqa: E402

CH = load_changes(CONTENT.read_text(encoding='utf-8'))
if CH['rows']:
    assert f'({CH["latest"]})' in hp.text, f'header version must match latest CHANGES edition {CH["latest"]}'
    page_break()  # 표지 마지막 문장 끝에 쪽 나누기
    p = doc.add_paragraph(style='Heading 1'); runs(p, f'{CH["base"].split()[0]} 대비 변경 요약')
    para_border(p, 'bottom', 8, INK, 3); p.style = doc.styles['Normal']; p.runs[0].bold = True; p.runs[0].font.size = Pt(10)  # 목차에 넣지 않는다
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(4)
    runs(p, f'기준판 {CH["base"]} → 최신판 {CH["latest"]}. 최신판 변경은 맨 위·굵게. 근거 열은 바뀐 수치의 근거 성격(공식 자료·언론 보도·업체 자료·직접 관측·당사 추정).',
         size=8.5, color=MUTED)
    t = doc.add_table(rows=1 + len(CH['rows']), cols=4)
    fix_table(t, fit([1.25, 1.0, 6.2, 1.45]))
    borders(t._tbl.tblPr, top=(12, NAVY), bottom=(12, NAVY), insideH=(4, RULE), left=None, right=None, insideV=None)
    for i, h in enumerate(('판', '장', '바뀐 점', '근거')):
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
    page_break()
    p = doc.add_paragraph(); runs(p, '검토 중 의문사항', size=10, bold=True)
    para_border(p, 'bottom', 8, INK, 3); p.paragraph_format.left_indent = Cm(0)
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
PHOTOS = Path(__file__).resolve().parent / 'photos'  # 제품 사진 묶음 그림 (tools/report/photos.py로 생성)
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
