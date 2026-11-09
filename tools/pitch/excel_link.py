"""엑셀 시뮬레이션 읽기: 연차별 값을 돌려준다.

파일이 엑셀에 열려 있으면 엑셀이 계산한 값을 그대로 읽는다(모든 함수 지원).
닫혀 있으면 아래 방식으로 수식을 직접 계산한다.

수기로 만든 시트라 계산 결과가 파일에 저장돼 있지 않은 경우가 있어,
openpyxl로 수식을 읽고 여기서 계산한다(사칙연산, %, SUM/MIN/MAX/ROUND/AVERAGE 지원).
행은 B열 라벨로 찾으므로 행 위치가 바뀌어도 라벨만 같으면 된다.
"""
import re
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

ROOT = Path(__file__).resolve().parents[2]
SIM_DIR = ROOT / '80-산출물' / '시뮬레이션모델링'
SECTIONS = ('본체', '패치')  # '수량'처럼 겹치는 라벨은 '본체 수량'·'패치 수량'으로 구분


def latest_workbook():
    def ver(p):
        m = re.search(r'v(\d+)\.xlsx$', p.name)
        return int(m.group(1)) if m else -1
    files = [p for p in SIM_DIR.glob('*.xlsx') if not p.name.startswith('~$')]
    return max(files, key=ver)


def excel_values(path, sheet=None):
    """엑셀에서 이 파일이 열려 있으면 엑셀이 계산한 값을 {'C5': 값} 으로 돌려준다(없으면 None).
    엑셀이 직접 계산하므로 IF·VLOOKUP 같은 함수도 그대로 반영된다."""
    try:
        import pythoncom
        import win32com.client as win32
        pythoncom.CoInitialize()
        app = win32.GetActiveObject('Excel.Application')
    except Exception:
        return None
    target = str(Path(path).resolve()).lower()
    for i in range(1, app.Workbooks.Count + 1):
        wb = app.Workbooks(i)
        if wb.FullName.lower() != target:
            continue
        used = wb.Worksheets(sheet or 1).UsedRange
        r0, c0 = used.Row, used.Column
        grid = used.Value2
        if not isinstance(grid, tuple):
            grid = ((grid,),)
        out = {}
        for ri, row in enumerate(grid):
            for ci, v in enumerate(row):
                if isinstance(v, float) and v.is_integer():
                    v = int(v)
                out[f'{get_column_letter(c0 + ci)}{r0 + ri}'] = v
        return out
    return None


class Sheet:
    def __init__(self, ws, live=None):
        self.ws = ws
        self.live = live  # 엑셀이 계산한 값(파일이 엑셀에 열려 있을 때)
        self.cache = {}

    def value(self, ref):
        ref = ref.replace('$', '').upper()
        if ref in self.cache:
            return self.cache[ref]
        if self.live is not None:
            v = self.live.get(ref)
            self.cache[ref] = 0 if v is None else v
            return self.cache[ref]
        v = self.ws[ref].value
        if isinstance(v, str) and v.startswith('='):
            v = self._eval(v[1:])
        if v is None:
            v = 0
        self.cache[ref] = v
        return v

    def _range(self, a, b):
        ca, ra = re.match(r'([A-Z]+)(\d+)', a).groups()
        cb, rb = re.match(r'([A-Z]+)(\d+)', b).groups()
        out = []
        for c in range(column_index_from_string(ca), column_index_from_string(cb) + 1):
            for r in range(int(ra), int(rb) + 1):
                v = self.value(f'{get_column_letter(c)}{r}')
                if isinstance(v, (int, float)):
                    out.append(v)
        return out

    def _eval(self, expr):
        expr = expr.replace('$', '')
        expr = re.sub(r'([A-Z]+\d+):([A-Z]+\d+)', lambda m: f'_R("{m.group(1)}","{m.group(2)}")', expr)
        expr = re.sub(r'(\d+(?:\.\d+)?)%', r'(\1/100)', expr)
        expr = re.sub(r'(?<![A-Z"_])([A-Z]{1,3}\d+)(?![\d"(])', lambda m: f'_V("{m.group(1)}")', expr)
        flat = lambda args: [x for a in args for x in (a if isinstance(a, list) else [a])]
        env = {
            '_V': self.value, '_R': self._range,
            'SUM': lambda *a: sum(flat(a)), 'MIN': lambda *a: min(flat(a)), 'MAX': lambda *a: max(flat(a)),
            'AVERAGE': lambda *a: sum(flat(a)) / len(flat(a)), 'ROUND': lambda x, n=0: round(x, int(n)),
        }
        try:
            return eval(expr, {'__builtins__': {}}, env)
        except Exception as e:  # 지원하지 않는 함수 등
            raise ValueError(f'엑셀 수식 계산 실패: ={expr} ({e})')


def load(path=None):
    """{'file': 경로, 'years': [1,2,3], 'rows': {라벨: [연차별 값]}, 'notes': [...]} 반환."""
    path = Path(path) if path else latest_workbook()
    wb = openpyxl.load_workbook(path)
    # 회의용 시트(excel_meeting.py)가 있으면 그 시트를 읽는다
    ws = wb['해커톤 회의'] if '해커톤 회의' in wb.sheetnames else wb.worksheets[0]
    sh = Sheet(ws, excel_values(path, ws.title))
    # 연차 헤더 행: C열부터 1,2,3... 이 있는 첫 행
    year_row, cols = None, []
    for r in range(1, ws.max_row + 1):
        vals = [(c, ws.cell(r, c).value) for c in range(3, ws.max_column + 1)]
        vals = [(c, v) for c, v in vals if v is not None]  # 빈 칸은 건너뜀(0으로 세지 않음)
        nums = [(c, v) for c, v in vals if isinstance(v, int) and not isinstance(v, bool)]
        if len(nums) != len(vals):
            continue
        if len(nums) >= 2 and [v for _, v in nums] == list(range(1, len(nums) + 1)):
            year_row, cols = r, [c for c, _ in nums]
            break
    if not year_row:
        raise ValueError('연차 헤더(1, 2, 3 …) 행을 찾지 못함')
    rows, section, notes = {}, None, []
    for r in range(year_row + 1, ws.max_row + 1):
        label = ws.cell(r, 2).value
        if not isinstance(label, str) or not label.strip():
            continue
        label = label.strip()
        if label in SECTIONS:
            section = label
        key = f'{section} {label}' if label == '수량' and section else label
        rows[key] = [sh.value(f'{get_column_letter(c)}{r}') for c in cols]
    for r in range(1, year_row):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                notes.append(v)
    return {'file': path, 'sheet': ws.title, 'years': list(range(1, len(cols) + 1)), 'rows': rows, 'notes': notes}


if __name__ == '__main__':
    import sys
    d = load(sys.argv[1] if len(sys.argv) > 1 else None)
    print(d['file'].name, d['sheet'], d['notes'])
    for k, v in d['rows'].items():
        print(f'{k:10s}', [round(x, 3) if isinstance(x, float) else x for x in v])
