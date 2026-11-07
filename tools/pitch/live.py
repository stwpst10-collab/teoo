"""실시간 모드: 사용자가 워드를 열어 둔 채로 같은 문서를 고친다(워드 자동화).

  python tools/pitch/live.py open                       문서를 워드로 연다(이미 열려 있으면 그대로)
  python tools/pitch/live.py add <차트ID> --after "문구"  문구가 있는 문단 뒤에 차트 추가
  python tools/pitch/live.py sync                       엑셀 최신본 값으로 연동 차트·표·출처 문구 갱신
  python tools/pitch/live.py footnotes                  본문 [n] → 쪽 아래 각주
  python tools/pitch/live.py save                       저장만

- 사용자가 쓰는 중인 글(저장 전 포함)은 건드리지 않고, 작업마다 저장한다.
- 작업 하나는 워드 '실행 취소' 한 번(Ctrl+Z)으로 되돌릴 수 있게 묶는다.
"""
import argparse
import json
import re
import sys
import tempfile
from pathlib import Path

import pythoncom
import win32com.client as win32

sys.path.insert(0, str(Path(__file__).resolve().parent))
import charts  # noqa: E402
import doc as d  # noqa: E402
import excel_link  # noqa: E402

CM = 28.3465
MUTED_BGR = 0x5C5F5F  # 워드 색은 BGR
TMP = Path(tempfile.gettempdir()) / 'pitch_live'
TMP.mkdir(exist_ok=True)


def word_doc(open_if_needed=True):
    target = str(d.DOC.resolve()).lower()
    try:
        app = win32.GetActiveObject('Word.Application')
    except pythoncom.com_error:
        if not open_if_needed:
            sys.exit('워드가 실행 중이 아님')
        app = win32.Dispatch('Word.Application')
    for i in range(1, app.Documents.Count + 1):
        doc = app.Documents(i)
        if doc.FullName.lower() == target:
            return app, doc
    if not open_if_needed:
        sys.exit('문서가 열려 있지 않음')
    app.Visible = not __import__('os').environ.get('PITCH_DOC')  # 시험용 경로면 창을 띄우지 않음
    return app, app.Documents.Open(str(d.DOC.resolve()))


class Undo:
    def __init__(self, app, name):
        self.app, self.name = app, name

    def __enter__(self):
        self.app.ScreenUpdating = False
        self.app.UndoRecord.StartCustomRecord(f'Claude: {self.name}')

    def __exit__(self, *exc):
        self.app.UndoRecord.EndCustomRecord()
        self.app.ScreenUpdating = True


def _png(cid, xl):
    p = TMP / f'{cid}.png'
    p.write_bytes(charts.render(cid, xl))
    return str(p)


def _new_para_after(para):
    para.Range.InsertParagraphAfter()
    new = para.Next(1)
    new.Range.Style = -1  # 표준
    new.Range.ParagraphFormat.LeftIndent = 0
    new.Range.ParagraphFormat.FirstLineIndent = 0
    try:
        new.Range.ListFormat.RemoveNumbers()
    except Exception:
        pass
    return new


def _caption(para, text):
    rng = para.Range
    rng.MoveEnd(1, -1)
    rng.Text = text
    rng.Font.Name = d.FONT
    rng.Font.Size = 8.5
    rng.Font.Bold = False
    rng.Font.Color = MUTED_BGR
    para.Format.SpaceAfter = 10


def _picture_into(para, cid, xl):
    rng = para.Range
    rng.Collapse(1)
    shp = para.Range.InlineShapes.AddPicture(_png(cid, xl), False, True, rng)
    shp.LockAspectRatio = True
    shp.Width = 16 * CM
    shp.AlternativeText = f'chart:{cid}'
    shp.Title = charts.CHARTS[cid]['title']
    para.Alignment = 1
    para.KeepWithNext = True
    return shp


def add(cid, after, source=None):
    if cid not in charts.CHARTS:
        sys.exit(f'없는 차트 ID: {cid}. 등록된 ID: {", ".join(charts.CHARTS)}')
    app, doc = word_doc()
    rng = doc.Content
    if not rng.Find.Execute(FindText=after[:250], MatchCase=False, Forward=True, Wrap=0):
        sys.exit(f'"{after}" 문구를 찾지 못함')
    xl = excel_link.load() if charts.CHARTS[cid]['excel'] else None
    with Undo(app, f'차트 {cid}'):
        anchor = rng.Paragraphs(1)
        pic = _new_para_after(anchor)
        _picture_into(pic, cid, xl)
        cap = _new_para_after(pic)
        cap.Alignment = 0
        cap.KeepWithNext = False
        _caption(cap, d.xl_caption_text(xl) if xl else (source or '자료: 그날의 류'))
    doc.Save()
    print(f'추가: {cid} → "{anchor.Range.Text[:30].strip()}…" 뒤, 저장함')


def sync():
    app, doc = word_doc()
    xl = excel_link.load()
    done = []
    with Undo(app, '엑셀 반영'):
        shapes = [doc.InlineShapes(i) for i in range(1, doc.InlineShapes.Count + 1)]
        for shp in shapes:
            alt = shp.AlternativeText or ''
            if not alt.startswith('chart:'):
                continue
            cid = alt[6:]
            if cid not in charts.CHARTS:
                print(f'경고: 등록되지 않은 차트 {cid}'); continue
            width, rng = shp.Width, shp.Range
            shp.Delete()
            new = doc.InlineShapes.AddPicture(_png(cid, xl), False, True, rng)
            new.LockAspectRatio = True
            new.Width = width
            new.AlternativeText = alt
            new.Title = charts.CHARTS[cid]['title']
            done.append(cid)
        for ti in range(1, doc.Tables.Count + 1):
            t = doc.Tables(ti)
            if t.Title != 'excel:pl_table':
                continue
            rows = d.pl_rows(xl)
            while t.Rows.Count < len(rows):
                t.Rows.Add()
            while t.Rows.Count > len(rows):
                t.Rows(t.Rows.Count).Delete()
            for i, row in enumerate(rows, 1):
                for j, val in enumerate(row, 1):
                    t.Cell(i, j).Range.Text = val
            done.append('표:pl_table')
        for pi in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(pi)
            if p.Range.Text.startswith(d.XL_CAPTION):
                _caption(p, d.xl_caption_text(xl))
    doc.Save()
    print(f'갱신: {", ".join(done) or "없음"} (엑셀: {xl["file"].name}), 저장함')


def footnotes():
    app, doc = word_doc()
    store = Path(__file__).with_name('sources.json')
    sources = {int(k): v for k, v in json.loads(store.read_text(encoding='utf-8')).items()}
    made, missing = 0, set()
    with Undo(app, '각주'):
        fn_style = doc.Styles(-30)  # 각주 텍스트
        fn_style.Font.Name = d.FONT
        fn_style.Font.Size = 8
        fn_style.ParagraphFormat.LineSpacingRule = 0  # 1줄 간격(작업 기준 6장)
        fn_style.ParagraphFormat.SpaceBefore = 0
        fn_style.ParagraphFormat.SpaceAfter = 0
        # 한글판 워드에서 와일드카드 검색이 안 돼 번호별 일반 검색으로 찾는다
        found = sorted({int(m) for m in re.findall(r'\[(\d{1,2})\]', doc.Content.Text)})
        missing = {n for n in found if n not in sources}
        for n in found:
            if n in missing:
                continue
            while True:
                rng = doc.Content
                if not rng.Find.Execute(f'[{n}]'):
                    break
                before = rng.Duplicate
                before.SetRange(rng.Start - 1, rng.Start)
                if before.Text == ' ':
                    rng.SetRange(rng.Start - 1, rng.End)
                rng.Text = ''
                fn = doc.Footnotes.Add(Range=rng)
                fn.Range.Text = sources[n]
                made += 1
    doc.Save()
    print(f'각주 {made}개 생성' + (f', 출처 목록에 없는 번호: {sorted(missing)}' if missing else '') + ', 저장함')


def renumber():
    """장(제목 1, 'N. ' 시작) 아래 절(제목 2)에 N.1, N.2 … 번호. 절이 하나뿐인 장은 번호 없음. 목차 갱신."""
    app, doc = word_doc()
    prefix = re.compile(r'^\d+\.\d+\s+')
    chapters, cur = [], None
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        if p.Range.Information(12):
            continue
        t = p.Range.Text.strip()
        if p.OutlineLevel == 1 and re.match(r'^\d+\. ', t):
            cur = [int(t.split('.')[0]), []]
            chapters.append(cur)
        elif p.OutlineLevel == 2 and cur is not None:
            cur[1].append(p)
    changed = 0
    with Undo(app, '절 번호 정리'):
        for num, secs in chapters:
            for k, p in enumerate(secs, 1):
                r = p.Range
                m = prefix.match(r.Text)
                if m:
                    old = doc.Range(r.Start, r.Start + len(m.group(0)))
                    old.Delete()
                if len(secs) > 1:
                    doc.Range(p.Range.Start, p.Range.Start).InsertBefore(f'{num}.{k}  ')
                changed += 1
        for t in range(1, doc.TablesOfContents.Count + 1):
            doc.TablesOfContents(t).Update()
    doc.Save()
    single = [n for n, s in chapters if len(s) == 1]
    print(f'절 {changed}개 번호 정리, 번호 없앤 장(절 1개): {single}, 목차 갱신, 저장함')


def open_doc():
    app, doc = word_doc()
    app.Visible = True
    app.Activate()
    print(f'열림: {doc.Name}')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest='cmd', required=True)
    for c in ('open', 'sync', 'footnotes', 'save', 'renumber'):
        sub.add_parser(c)
    a = sub.add_parser('add'); a.add_argument('cid'); a.add_argument('--after', required=True); a.add_argument('--source')
    args = ap.parse_args()
    if args.cmd == 'add':
        add(args.cid, args.after, args.source)
    elif args.cmd == 'save':
        word_doc()[1].Save(); print('저장함')
    else:
        {'open': open_doc, 'sync': sync, 'footnotes': footnotes, 'renumber': renumber}[args.cmd]()
