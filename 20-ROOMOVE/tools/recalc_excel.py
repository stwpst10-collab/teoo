# Excel로 전체 재계산·저장. 'Office 인증 마법사' 창이 뜨면 자동으로 닫음
import sys, os, time, threading
import win32com.client as w, win32gui, win32process, win32con, pythoncom
path = os.path.abspath(sys.argv[1])
xl = w.DispatchEx('Excel.Application')
_, pid = win32process.GetWindowThreadProcessId(xl.Hwnd)
stop = False
def closer():
    while not stop:
        def cb(h, _):
            try:
                if win32process.GetWindowThreadProcessId(h)[1] == pid and win32gui.GetClassName(h) == 'NUIDialog':
                    win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0); print('closed dialog', flush=True)
            except Exception: pass
            return True
        win32gui.EnumWindows(cb, None); time.sleep(0.5)
threading.Thread(target=closer, daemon=True).start()
def rt(fn):
    for _ in range(240):
        try: return fn()
        except Exception: time.sleep(0.5)
    raise RuntimeError('excel busy')
try:
    rt(lambda: setattr(xl, 'DisplayAlerts', False))
    wb = rt(lambda: xl.Workbooks.Open(path))
    rt(lambda: xl.CalculateFull())
    while rt(lambda: xl.CalculationState) != 0: time.sleep(0.5)
    rt(lambda: wb.Save()); rt(lambda: wb.Close(False))
    print('saved')
finally:
    rt(lambda: xl.Quit()); stop = True
