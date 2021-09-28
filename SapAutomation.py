import win32gui
import pyautogui
import pywinauto
import time
import sys


win_1 = "SAP Easy Access"
win_2 = "Backlog list"
win_3 = "ABAP: Variant Directory of Program ZWRUECKSTAND_MX"
delay = 2


def press_key(t, key, count):
    time.sleep(t)
    pyautogui.press(key, presses=count)


def text_input(t, text):
    time.sleep(t)
    pyautogui.typewrite(text)


def press_hotkey(t, args):
    time.sleep(t)
    pyautogui.hotkey(*args)


def wait_for_windows():
    try:
        app = pywinauto.application.Application().connect(path=r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")# x64
    except pywinauto.application.ProcessNotFoundError:
        try:
            app = pywinauto.application.Application().connect(path=r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe")# x32
        except pywinauto.application.ProcessNotFoundError:
            pyautogui.alert('Програму SAP не запущено !','Помилка !')
            sys.exit()
    dlg = app.window(title_re=".*Backlog.*")
    ctrl = dlg.child_window(class_name="SAPALVGridControl")
    return ctrl
    

def find_win(triger, window):
    while triger:
        whnd = win32gui.FindWindowEx(None, None, None, window)
        if not (whnd == 0):
            win32gui.SetForegroundWindow(whnd)
            triger = False


whnd = win32gui.FindWindowEx(None, None, None, win_1)
if (whnd == 0):
    pyautogui.alert('Головне вікно SAP не знайдено, залогуйтесь в систему !','Помилка !')
    sys.exit(0)


def sap_update():
    t = True
    find_win(t, win_1)
    press_key(delay, 'esc', 5)
    t = True
    find_win(t, win_1)
    text_input(delay, '/nzwrs_mx')
    t = True
    find_win(t, win_1)
    press_key(delay, 'enter', 1)
    t = True
    find_win(t, win_2)
    press_hotkey(delay, ['shift', 'f5'])
    t = True
    find_win(t, win_3) 
    press_key(delay, 'right', 3) 
    t = True
    find_win(t, win_3)       
    press_key(delay, 'f2', 1)
    t = True
    find_win(t, win_2)
    press_key(delay, 'f8', 1)

def reloads():
    whnd = win32gui.FindWindowEx(None, None, None, win_2)
    if not (whnd == 0):
        win32gui.SetForegroundWindow(whnd)
    press_key(delay, 'f8', 1)
