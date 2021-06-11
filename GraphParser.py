import csv
import ctypes.wintypes
import re
import os


def documents_path():
    CSIDL_PERSONAL = 5       # My Documents
    SHGFP_TYPE_CURRENT = 0   # Get current, not default value

    buf= ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)

    path = buf.value + r'\SAP\SAP GUI'
    return path


def run_update():
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
                tmp = row[1].split("|")[3]
                tmp = int(tmp)
                try:
                    if int(tmp) > 2:
                        for i in row:
                            try:
                                if re.search(r'seat', i.split('|')[4].lower()) and not re.search(r'bmw', i.split('|')[4].lower()):
                                    seat += int(i.split('|')[5])
                                if re.search(r'bmw', i.split('|')[4].lower()):
                                    bmw += int(i.split('|')[5])
                                if re.search(r'audi', i.split('|')[4].lower()):
                                    audi += int(i.split('|')[5])
                                if re.search(r'br223', i.split('|')[4].lower()):
                                    br223 += int(i.split('|')[5])
                            except IndexError:
                                continue
                except ValueError:
                    continue                    
            except IndexError:
                continue
            except ValueError:
                continue
    return seat, bmw, audi, br223            
