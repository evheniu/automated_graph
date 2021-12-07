import csv
import ctypes.wintypes
import re
import os


def documents_path():
    CSIDL_PERSONAL = 5       # My Documents
    SHGFP_TYPE_CURRENT = 0   # Get current, not default value

    buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)

    path = buf.value + r'\SAP\SAP GUI'
    return path


def run_update():
    sk38 = 0
    seat = 0    
    bmw = 0    
    audi = 0    
    br223 = 0    
    tmp = 0
    if os.path.isfile(documents_path() + r'\data.txt'):
        pass
    else:
        with open(documents_path() + r'\data.txt', 'w') as f:
            f.close()
    wb =  documents_path() + r'\data.txt'
    with open(wb, newline='') as File:  
        reader = csv.reader(File)
        for row in reader:
            try:
                tmp = row[0].split("|")[4]
                tmp = int(tmp)
                try:
                    if int(tmp) > 1:
                        for i in row:
                            try:
                                if re.search(r'seat', i.split('|')[2].lower()) and not re.search(r'bmw', i.split('|')[2].lower()):
                                    seat += float(i.split('|')[3])
                                if re.search(r'bmw', i.split('|')[2].lower()):
                                    bmw += float(i.split('|')[3])
                                if re.search(r'br223', i.split('|')[2].lower()):
                                    br223 += float(i.split('|')[3])
                                if re.search(r'sk38', i.split('|')[2].lower()):
                                    sk38 += float(i.split('|')[3])
                            except IndexError:
                                continue
                    if int(tmp) > 2:
                        for i in row:
                            try:
                                if re.search(r'audi', i.split('|')[2].lower()):
                                    audi += float(i.split('|')[3])
                            except IndexError:
                                continue                            
                except ValueError:
                    continue                    
            except IndexError:
                continue
            except ValueError:
                continue
    return seat, bmw, audi, br223, sk38            
