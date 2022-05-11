import win32gui
import win32process
import win32api
import time
from systray import SystemTray

vbaClientWhDl = 0


def main():
    win32gui.EnumWindows(callback, None)

    print("ts:")
    print(vbaClientWhDl)
    # fg_win = win32gui.GetForegroundWindow()
    fg_thread, fg_process = win32process.GetWindowThreadProcessId(vbaClientWhDl)
    current_thread = win32api.GetCurrentThreadId()
    win32process.AttachThreadInput(current_thread, fg_thread, True)

    try:
        print(vbaClientWhDl)
        # win32gui.SetFocus(vbaClientWhDl)
        win32gui.SetActiveWindow(vbaClientWhDl)
        print(win32gui.GetCaretPos())
    finally:
        win32process.AttachThreadInput(current_thread, fg_thread, False)  # detach

    time.sleep(40)


def callback(hwnd, extra):
    global vbaClientWhDl
    if "Microsoft Visual Basic" in win32gui.GetWindowText(hwnd):
        print(f"window text: '{win32gui.GetWindowText(hwnd)}'")
        vbaClientWhDl = hwnd


main()
