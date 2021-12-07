import subprocess
import time
import pyautogui
import win32gui

# file with parameters
with open('login_info.txt') as f:
    lines = f.readlines()

# first line login SAP, 
# second line password SAP, 
# third line destination tcode
USER_ID = lines[0]
PASSWORD_SAP = lines[1]
TCODE = lines[2]
#
win_title_1 = "Backlog list"
# input parameters
SAPshortcut_Program_parameters = [
        '/max',
        '-command=' + TCODE,
        '-language=EN', 
        '-system=BPR', 
        '-client=010', 
        '-user=' + USER_ID, 
        '-pw=' + PASSWORD_SAP]

def sign_in():
    try:
        subprocess.check_call([r"C:\Program Files\SAP\FrontEnd\SAPgui\sapshcut.exe", *SAPshortcut_Program_parameters])
    except FileNotFoundError:
        subprocess.check_call([r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe", *SAPshortcut_Program_parameters])
    
#wait when main window will be started
def check_window_exist():
    while True:
        whnd = win32gui.FindWindowEx(None, None, None, win_title_1)
        if whnd == 0:
            whnd = win32gui.FindWindowEx(None, None, None, win_title_1)
        else:
            time.sleep(3)
            pyautogui.press('ctrl', presses=10)
            break
