"""빌더가 만든 docx를 워드로 마무리한다: 목차·필드 갱신, 표 머리만 쪽 끝에 남은 곳 정리, docx·pdf 저장.

    python tools/report/finish_word.py <빌드.docx> <최종.docx> <최종.pdf>

- 사용자가 열어 둔 워드는 건드리지 않는다. 항상 새 워드 인스턴스(DispatchEx)를 띄워 숨긴 채 쓰고, 끝나면 그 인스턴스만 종료한다
  (작업 기준 7장, 2026-11-07 "워드 왜 자꾸 종료되냐").
- '다음 문단과 함께'를 쓰지 않으므로(작업 기준 6장) 표 제목·머리말 행만 쪽 끝에 남을 수 있다. 쪽을 매긴 뒤
  그런 표를 찾아 바로 앞 문장 끝에 쪽 나누기를 넣는다(앞이 표·그림이면 1pt 높이 줄).
- 끝나면 쪽 수, 각주 수, 빈 쪽 여부를 출력한다.
"""
import sys

import win32com.client as wc

WD_PAGE = 3  # wdActiveEndPageNumber
WD_PAGE_BREAK = 7


def _page(rng):
    return rng.Information(WD_PAGE)


def fix_orphan_table_heads(d, max_rounds=10):
    """표 제목 행(1)·머리말 행(2)과 첫 데이터 행(3)이 다른 쪽이면 표 앞에서 쪽을 나눈다."""
    fixed = []
    for _ in range(max_rounds):
        d.Repaginate()
        target = None
        for i in range(1, d.Tables.Count + 1):
            t = d.Tables(i)
            if t.Rows.Count < 3 or t.NestingLevel > 1:
                continue
            if _page(t.Rows(1).Range) != _page(t.Rows(3).Range):
                target = (i, t)
                break
        if not target:
            return fixed
        i, t = target
        start = t.Range.Start
        prev = d.Range(start - 1, start - 1).Paragraphs(1)  # 표 바로 앞 문단
        text = prev.Range.Text.strip()
        if text and not prev.Range.Information(12) and prev.Range.InlineShapes.Count == 0:
            at = d.Range(prev.Range.End - 1, prev.Range.End - 1)  # 문단 기호 바로 앞 = 문장 끝
            at.InsertBreak(WD_PAGE_BREAK)
        else:
            at = d.Range(start, start)
            at.InsertParagraphBefore()
            p = d.Range(start, start).Paragraphs(1)
            p.Range.Font.Size = 1
            p.SpaceBefore = p.SpaceAfter = 0
            p.LineSpacingRule = 4; p.LineSpacing = 1  # 고정 1pt
            d.Range(p.Range.End - 1, p.Range.End - 1).InsertBreak(WD_PAGE_BREAK)
        fixed.append(i)
    return fixed


def blank_pages(d):
    """본문 글자가 없는 쪽 번호(머리말·꼬리말 제외)."""
    n = d.ComputeStatistics(2)
    empty = []
    for pg in range(1, n + 1):
        r = d.GoTo(1, 1, pg)  # wdGoToPage, wdGoToAbsolute
        r2 = d.GoTo(1, 1, pg + 1) if pg < n else d.Content
        end = r2.Start if pg < n else d.Content.End
        body = d.Range(r.Start, end)
        if not body.Text.strip() and body.InlineShapes.Count == 0 and body.Tables.Count == 0:
            empty.append(pg)
    return empty


def main(src, out_docx, out_pdf):
    app = wc.DispatchEx('Word.Application')  # 사용자 워드와 분리된 새 인스턴스
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        d = app.Documents.Open(src, ReadOnly=False, AddToRecentFiles=False)
        d.Fields.Update()
        for i in range(1, d.TablesOfContents.Count + 1):
            d.TablesOfContents(i).Update()
        fixed = fix_orphan_table_heads(d)
        for i in range(1, d.TablesOfContents.Count + 1):  # 쪽이 바뀌었으니 목차 쪽 번호 다시
            d.TablesOfContents(i).Update()
        d.Repaginate()
        print('pages', d.ComputeStatistics(2), 'footnotes', d.Footnotes.Count,
              'table-head fixes', fixed, 'blank pages', blank_pages(d))
        d.SaveAs2(out_docx, 16)
        d.ExportAsFixedFormat(out_pdf, 17)
        d.Close(False)
    finally:
        app.Quit()


if __name__ == '__main__':
    main(*sys.argv[1:4])
