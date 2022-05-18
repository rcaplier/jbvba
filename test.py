import array
import math
import time
from xmlrpc.client import Boolean

import keyboard
import win32api
import win32clipboard
import win32con
import win32gui
import win32process
from infi.systray import SysTrayIcon

hover_text = "JBVBA"
vbaClientWhDl = -1
projectWhDl = -1  # handle to the let project docker todo
codeWhDl = -1

quitApp = False


def bye(sysTrayIcon):
    global quitApp
    quitApp = True
    print('Bye, then.')


def cbGetCodeWin(hwnd, param):
    global codeWhDl
    if param in win32gui.GetWindowText(hwnd):
        codeWhDl = hwnd


# Callback to get VBA windows handle AND code window handle (only work if code window is maximised for now)
def cbGetVBAWin(hwnd, extra):
    global vbaClientWhDl, codeWhDl
    if "Microsoft Visual Basic for Applications" in win32gui.GetWindowText(hwnd):
        vbaClientWhDl = hwnd
        main_window_title = win32gui.GetWindowText(hwnd)

        # if code window is maximised, then it will be printed in the main window title between brackets
        if "[" in main_window_title:
            # We have to clean it first (remove the brackets around)
            code_window_title = str(main_window_title.split("[")[1]).strip()
            if "]" in code_window_title:
                code_window_title = code_window_title.replace("]", "", 1)

            win32gui.EnumChildWindows(vbaClientWhDl, cbGetCodeWin, code_window_title)


# Method to check if VBA editor window is foreground ( todo : and has focus)
def checkIfVBAIDEHasFocus() -> Boolean:
    win32gui.EnumWindows(cbGetVBAWin, None)
    global vbaClientWhDl
    fg_win = win32gui.GetForegroundWindow()
    if vbaClientWhDl == fg_win:
        sysTrayIcon.update('ok.ico')
        return True
    else:
        sysTrayIcon.update('no.ico')
        return False


# Get the current position of the caret in the VBA window
# It return a tuple -> ( x, y ) where x is the column and y is the line.
# 34 is the first column and each column is 8 wide : 34 + 8x
# 30 is the first line and each line is 16 high : 30 + 16y
def getCaretPosition() -> tuple:
    caret_position = (0, 0)
    fg_thread, fg_process = win32process.GetWindowThreadProcessId(vbaClientWhDl)
    current_thread = win32api.GetCurrentThreadId()
    win32process.AttachThreadInput(current_thread, fg_thread, True)
    try:
        win32gui.SetActiveWindow(vbaClientWhDl)
        caret_position = win32gui.GetCaretPos()
    finally:
        win32process.AttachThreadInput(current_thread, fg_thread, False)  # detach
        return caret_position


def getCaretColumn() -> int:
    if checkIfVBAIDEHasFocus():
        # print("Column caret : " + str(getCaretPosition()[0]))
        return getCaretPosition()[0]


def getCaretLine() -> int:
    if checkIfVBAIDEHasFocus():
        # print("Line caret : " + str(getCaretPosition()[1]))
        return getCaretPosition()[1]


def isCaretAtBeginningOfLine() -> Boolean:
    if checkIfVBAIDEHasFocus():
        if getCaretColumn() == 34:
            return True
    return False


def getLineNumber() -> int:
    if checkIfVBAIDEHasFocus():
        caret_line = getCaretLine()
        if caret_line == 30:
            print("Line : 1")
            return 1
        if caret_line > 30:
            print("Line : " + str(math.ceil((caret_line - 30) / 16) + 1))
            return math.ceil((caret_line - 30) / 16) + 1
        return -1


def getColumnNumber() -> int:
    if checkIfVBAIDEHasFocus():
        caret_column = getCaretColumn()
        if caret_column == 34:
            print("Column : 1")
            return 1
        if caret_column > 34:
            print("Column : " + str(math.ceil((caret_column - 34) / 8) + 1))
            return math.ceil((caret_column - 34) / 8) + 1
        return -1


def getCaretAtBegin():
    if checkIfVBAIDEHasFocus():
        while not isCaretAtBeginningOfLine():
            keyboard.send('home')


def commentLine():
    if checkIfVBAIDEHasFocus():
        getCaretAtBegin()
        keyboard.send("'")
        keyboard.send('end')


def test4cb(hwnd, param):
    print("child")
    print(str(hwnd))


def GetCaretWindowText(hWndCaret, Selected=False):  # As String

    startpos = 0
    endpos = 0

    txt = ""

    if hWndCaret:

        buf_size = 1 + win32gui.SendMessage(hWndCaret, win32con.WM_GETTEXTLENGTH, 0, 0)
        if buf_size:
            buffer = win32gui.PyMakeBuffer(buf_size)
            win32gui.SendMessage(hWndCaret, win32con.WM_GETTEXT, buf_size, buffer)
            address, length = win32gui.PyGetBufferAddressAndLen(buffer)
            txt = win32gui.PyGetString(address, length)
            # txt = buffer[:buf_size]

        if Selected and buf_size:
            selinfo = win32gui.SendMessage(hWndCaret, win32con.EM_GETSEL, 0, 0)
            endpos = win32api.HIWORD(selinfo)
            startpos = win32api.LOWORD(selinfo)
            return txt[startpos: endpos]

    return txt


def test3cb(hwnd, param):
    print(win32gui.GetClassName(hwnd))
    print(GetCaretWindowText(hwnd))
    if win32gui.GetClassName(hwnd) == "ObtbarWndClass":
        code_pane = hwnd

        # test = win32gui.GetDlgCtrlID(hwnd)
        # print(test)

        # test = win32gui.GetDlgItemText(hwnd, test)

        # test = win32gui.SendMessage(test, win32con.CB_GETLBTEXTLEN, 0, 0)
        # print(str(test))
        cb_len = win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
        print("Text length : " + str(cb_len))
        # buffer = win32gui.PyMakeBuffer(cb_len)
        # test = win32gui.SendMessage(hwnd, win32con.CB_GETLBTEXT, 0, buffer)
        # print(str(test))
        # text = buffer
        # print(text)

        buffer = win32gui.PyMakeBuffer(cb_len)

        win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, cb_len, buffer)  # read the text

        address, length = win32gui.PyGetBufferAddressAndLen(buffer)

        text = win32gui.PyGetString(address, length)

        print("text: " + text)
    return True


def getProjectCB(hwnd, param):
    global projectWhdl
    if "Project - " in win32gui.GetWindowText(hwnd):
        projectWhdl = hwnd


def getSelection():
    if checkIfVBAIDEHasFocus():
        # resp = win32gui.SendMessage(codeWhDl, win32con.WM_COPY, None, None)
        # print(str(resp))

        test_var = win32gui.GetClassName(codeWhDl)
        print(str(test_var))

        win32gui.EnumChildWindows(codeWhDl, test3cb, None)

        # win32gui.GetDlgItem(projectWhdl)

        # test_var2 = str(win32gui.GetFocus())
        # print(vbaClientWhDl)
        # print(test_var2)

        # dc = win32gui.GetWindowDC(vbaClientWhDl)
        # print(str(dc))

        # keyboard.press('ctrl')
        # keyboard.press('c')
        # keyboard.release('c')
        # keyboard.release('ctrl')

        win32clipboard.OpenClipboard(vbaClientWhDl)
        # win32clipboard.EmptyClipboard()
        # keyboard.press('ctrl')
        # keyboard.press('c')

        # keyboard.release('ctrl')
        # keyboard.release('c')
        # print(str(win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)))
        win32clipboard.CloseClipboard()


def start():
    keyboard.add_hotkey('ctrl+/', commentLine, suppress=True)
    keyboard.add_hotkey('ctrl+*', getSelection, suppress=True)

    # start keylogger
    # keyboard.on_press(callback=callback)

    while not quitApp:
        checkIfVBAIDEHasFocus()
        time.sleep(0.2)


def callback(event):
    """
    This callback is invoked whenever a keyboard event is occured
    (i.e when a key is released in this example)
    """
    # name = event.name
    # if len(name) > 0:
    #     print("  ")
    #     print(name)
    #     print("  ")


# menu_options = (('Say Hello', "hello.ico", hello),
#                 ('Do nothing', None, do_nothing),
#                 ('A sub-menu', "submenu.ico", (('Say Hello to Simon', "simon.ico", simon),
#                                                ('Do nothing', None, do_nothing),
#                                               ))
#                )
# sysTrayIcon = SysTrayIcon("py.ico", hover_text, menu_options, on_quit=bye, default_menu_index=1)
# TODO : Clique gauche = donne le focus à la window VBA
sysTrayIcon = SysTrayIcon("ok.ico", hover_text, on_quit=bye, default_menu_index=1)
sysTrayIcon.start()

start()
