from xmlrpc.client import Boolean
from infi.systray import SysTrayIcon
import keyboard
import win32gui
import win32process
import win32api
import win32clipboard
import time
import math

hover_text = "JBVBA"
vbaClientWhDl = -1

quitApp = False


def bye(sysTrayIcon):
    global quitApp
    quitApp = True
    print('Bye, then.')


# Callback to get VBA windows handle
def cbGetVBAWin(hwnd, extra):
    global vbaClientWhDl
    if "Microsoft Visual Basic" in win32gui.GetWindowText(hwnd):
        vbaClientWhDl = hwnd


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


def getSelection():
    test_var = ""
    if checkIfVBAIDEHasFocus():
        win32clipboard.OpenClipboard(vbaClientWhDl)
        win32clipboard.EmptyClipboard()
        keyboard.press('ctrl')
        keyboard.press('c')
        #todo here
        keyboard.release('ctrl')
        keyboard.release('c')
        print(win32clipboard.GetClipboardData(1))
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
sysTrayIcon = SysTrayIcon("ok.ico", hover_text, on_quit=bye, default_menu_index=1)
sysTrayIcon.start()

start()
