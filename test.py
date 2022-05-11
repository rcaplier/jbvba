from xmlrpc.client import Boolean
from infi.systray import SysTrayIcon
import keyboard
import win32gui
import win32process
import win32api
import time
from pynput.keyboard import Key, Controller

inputKey = Controller()
hover_text = "JBVBA"
vbaClientWhDl = -1
quitApp = False


def bye(sysTrayIcon):
    global quitApp
    keyboard.stop_recording()
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


cpt = 0


def commentLine():
    global cpt
    if checkIfVBAIDEHasFocus():
        cpt = cpt + 1
        print("ok " + str(cpt))
        inputKey.tap(Key.home)
        # inputKey.press("'")
        # inputKey.press(Key.end)


def start():
    keyboard.add_hotkey('ctrl+/', commentLine, timeout=40000)

    # start the keylogger
    # keyboard.on_press(callback=callback)


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

while not quitApp:
    checkIfVBAIDEHasFocus()
    time.sleep(0.2)
