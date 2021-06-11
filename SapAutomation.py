import win32gui
import pyautogui
import pywinauto
import time
import sys


win_1 = "SAP Easy Access"
win_2 = "Backlog list"
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
    text_input(delay, '/nzwrs_mx')
    press_key(delay, 'enter', 1)
    t = True
    find_win(t, win_2)
    press_hotkey(delay, ['ctrl', 'a'])        
    text_input(delay, '0313')
    press_key(delay, 'down', 3)
    text_input(delay, '4000000')
    press_key(delay, 'tab', 1)
    text_input(delay, '4999999')
    press_key(delay, 'down', 9)
    press_key(delay, 'tab', 4)
    press_key(delay, 'down', 1)
    press_key(delay, 'f8', 1)
    t = True
    ctrl = wait_for_windows()
    while t:
        try:
            ctrl.parent
            t = False
        except pywinauto.findwindows.ElementNotFoundError:
            wait_for_windows()
    time.sleep(delay)
    press_hotkey(delay, ['ctrl', 'shift', 'f9'])
    press_key(delay, 'enter', 1)
    press_hotkey(delay, ['ctrl', 'a'])
    text_input(delay, 'data.txt')
    press_hotkey(delay, ['ctrl', 's'])
    press_key(delay, 'esc', 10)
    press_key(delay, 'esc', 10)
