import win32gui
import pyautogui
import time


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


def find_win_and_set_foreground(window):
    while True:
        whnd = win32gui.FindWindowEx(None, None, None, window)
        if not (whnd == 0):
            time.sleep(3)
            win32gui.SetForegroundWindow(whnd)
            break


def sap_update():
    find_win_and_set_foreground(win_2)    
    press_hotkey(delay, ['shift', 'f5'])
    find_win_and_set_foreground(win_3) 
    press_key(delay, 'right', 3) 
    find_win_and_set_foreground(win_3)       
    press_key(delay, 'f2', 1)
    find_win_and_set_foreground(win_2)
    press_key(delay, 'f8', 1)

def reloads():
    whnd = win32gui.FindWindowEx(None, None, None, win_2)
    if not (whnd == 0):
        win32gui.SetForegroundWindow(whnd)
    press_key(delay, 'f8', 1)