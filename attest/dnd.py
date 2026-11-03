"""Перетаскивание файлов в окно (Windows, без сторонних библиотек).

Окну Tk включается приём файлов (DragAcceptFiles), а его оконная процедура
подменяется: сообщение WM_DROPFILES отдаётся в callback, остальное - Tk как раньше.
Работает с Vista, в 32- и 64-битном Python.
"""
import ctypes
import sys
from ctypes import wintypes

WM_DROPFILES = 0x0233
GWL_WNDPROC = -4

_keep = []  # ссылки на callback'и, иначе сборщик мусора их удалит


def enable(root, on_files):
    """on_files(list[str]) вызывается при сбросе файлов на окно root."""
    if sys.platform != "win32":
        return False
    user32 = ctypes.windll.user32
    shell32 = ctypes.windll.shell32
    is64 = ctypes.sizeof(ctypes.c_void_p) == 8
    LRESULT = ctypes.c_ssize_t
    WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
    set_proc = user32.SetWindowLongPtrW if is64 else user32.SetWindowLongW
    set_proc.restype = ctypes.c_void_p
    set_proc.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
    call_proc = user32.CallWindowProcW
    call_proc.restype = LRESULT
    call_proc.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    shell32.DragQueryFileW.argtypes = [ctypes.c_void_p, wintypes.UINT, ctypes.c_wchar_p, wintypes.UINT]
    shell32.DragFinish.argtypes = [ctypes.c_void_p]
    shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]

    root.update_idletasks()
    hwnd = int(root.wm_frame(), 16)
    old = [None]
    dropped = []

    def poll():
        while dropped:
            on_files(dropped.pop(0))
        root.after(150, poll)

    def proc(h, msg, wp, lp):
        if msg == WM_DROPFILES:
            files = []
            for i in range(shell32.DragQueryFileW(wp, 0xFFFFFFFF, None, 0)):
                n = shell32.DragQueryFileW(wp, i, None, 0) + 1
                buf = ctypes.create_unicode_buffer(n)
                shell32.DragQueryFileW(wp, i, buf, n)
                files.append(buf.value)
            shell32.DragFinish(wp)
            # из оконной процедуры Tk вызывать нельзя (идёт обработка события Tcl) -
            # файлы забирает таймер ниже
            dropped.append(files)
            return 0
        return call_proc(old[0], h, msg, wp, lp)

    cb = WNDPROC(proc)
    _keep.append(cb)
    old[0] = set_proc(hwnd, GWL_WNDPROC, ctypes.cast(cb, ctypes.c_void_p))
    shell32.DragAcceptFiles(hwnd, True)
    root.after(150, poll)
    return True
