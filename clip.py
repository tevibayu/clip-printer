import time
import threading
import clipboard
import base64
import platform
import subprocess
import tempfile
import os
import win32print
import random
import winreg
import shutil
import sys
import getpass
import win32api
from zpl import Label
from PIL import Image
from infi.systray import SysTrayIcon


def say_hello(systray):
    print("Hello World!")

def on_quit_callback(systray):
    USER_NAME = getpass.getuser()
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Run', 0, winreg.KEY_SET_VALUE);
    winreg.DeleteValue(key, 'ClipPrinter');
    winreg.CloseKey(key);
    print('The application has been stopped in next reboot.')
    sys.exit(0)
    
#menu_options = (("List Printer", None, say_hello),)
menu_options = ()
systray = SysTrayIcon("Printer.ico", "Clip Printer", menu_options, on_quit=on_quit_callback)
systray.start()

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

#GHOSTSCRIPT_PATH = "C:\\Program Files\\gs\\gs9.26\\bin\\gswin64c.exe"
#GSPRINT_PATH = "C:\\Program Files\\gs\\gs9.26\\bin\\gsprint.exe"

GHOSTSCRIPT_PATH = resource_path('bin\gswin64c.exe')
GSPRINT_PATH = resource_path('bin\gsprint.exe')

#print(GHOSTSCRIPT_PATH)
#print(GSPRINT_PATH)

PRINTER_ERROR_STATES = (
    win32print.PRINTER_STATUS_NO_TONER,
    win32print.PRINTER_STATUS_NOT_AVAILABLE,
    win32print.PRINTER_STATUS_OFFLINE,
    win32print.PRINTER_STATUS_OUT_OF_MEMORY,
    win32print.PRINTER_STATUS_OUTPUT_BIN_FULL,
    win32print.PRINTER_STATUS_PAGE_PUNT,
    win32print.PRINTER_STATUS_PAPER_JAM,
    win32print.PRINTER_STATUS_PAPER_OUT,
    win32print.PRINTER_STATUS_PAPER_PROBLEM,
)

def printer_errorneous_state(printer, error_states=PRINTER_ERROR_STATES):
    prn_opts = win32print.GetPrinter(printer)
    status_opts = prn_opts[18]
    for error_state in error_states:
        if status_opts & error_state:
            return error_state
        
    return 0


def set_interval(func, sec, RUNNING_THREAD = on_quit_callback):
    def func_wrapper():
        set_interval(func, sec)
        func()
        
        
    t = threading.Timer(sec, func_wrapper)
    t.start()
    return t

def isBase64(s):
    try:
        return base64.b64encode(base64.b64decode(s)).decode("utf-8") == s
    except Exception:
        return False

def getClip():
    text = clipboard.paste()
    text = text.strip()
    with open('auth.key', 'rb') as f:
        key = f.read()
    word = key.decode("utf-8")
    
    if word in text:
        clipboard.copy("")
        coded_string = text.replace(word, "")
        coded_string = coded_string.split("_%_")
        
        if isBase64(coded_string[1]):
            Ext = coded_string[0]
            allCondsBinary = (Ext != 'txt')
            allCondsSpecial = (Ext == 'zpl')
            if allCondsBinary:
                if allCondsSpecial:
                    Ext = 'txt';
                Doc = base64.b64decode(coded_string[1]).decode("utf-8")
                print(Doc)
            else:
                Doc = base64.b64decode(coded_string[1]).decode("utf-8")
                
            printStatus(Doc, Ext)
        else:
            print("Format doesn't support")
            
def printStatus(Doc, Ext):
    printer_name = win32print.GetDefaultPrinter()
    prn = win32print.OpenPrinter(printer_name)
    error = printer_errorneous_state(prn)
    if error:
        print("ERROR occurred: ", error)
    else:
        print("Printer OK...")
        if Doc is not None:
            _platform = platform.system()
            printDoc(Doc, Ext, _platform)
        else:
            print("Document not found...")

    win32print.ClosePrinter(prn)
    
def printDoc(Doc, Ext, _platform):

    filename = tempfile.mktemp("." + Ext)
    print(filename)
    allCondsBinary = (Ext != 'txt')
    if allCondsBinary:
        open (filename, "wb").write(Doc)
    else:
        open (filename, "w").write(Doc)
    

    allCondsConvert = (Ext == 'png' or Ext == 'jpg' or Ext == 'jpeg')
    if allCondsConvert:
        converted = imageToZpl(filename, 65, 60)
        filename = tempfile.mktemp(".txt")
        print(filename)
        open (filename, "w").write(converted)
        
    
    if _platform == "Linux":
        print('From Linux!')
        subprocess.Popen(['/usr/bin/lpr', filename])
    elif _platform == "MacOS":
        print('From MacOS!')
        lpr =  subprocess.Popen("/usr/bin/lpr", stdin=subprocess.PIPE)
        lpr.stdin.write(Doc)
        lpr.stdin.close()
    elif _platform == "Windows":
        print('From Windows!')
        if Ext == 'pdf':
            currentprinter = win32print.GetDefaultPrinter()
            win32api.ShellExecute(0, 'open', GSPRINT_PATH, '-ghostscript "'+GHOSTSCRIPT_PATH+'" -printer "'+currentprinter+'" "'+filename+'"', '.', 0)
        else:
            os.startfile(filename, "print")

def imageToZpl(Image_, Width_, Height_):
    l = Label(Width_, Height_)
    height = 0

    height += 4
    image_width = (Width_-1)
    l.origin((l.width-image_width)/2, height)
    image_height = l.write_graphic(Image.open(Image_), image_width)
    l.endorigin()
    #open ('zplexport.zpl', "w").write(l.dumpZPL())
    return l.dumpZPL()

def copy_script():
    USER_NAME = getpass.getuser()
    src=sys.argv[0]
    
    dst = r'C:\Users\%s\AppData\Local\ClipPrinter' % USER_NAME
    if not os.path.isdir(dst):
        os.makedirs(dst)
    try:
        auth = r'C:\Users\%s\AppData\Local\ClipPrinter\auth.key' % USER_NAME
        if not os.path.isfile(auth):
            open (auth, "w").write('#your-token-application_')
        shutil.copy2(src,dst)
        
    except BaseException:
        pass
    
    dst = r'C:\Users\%s\AppData\Local\ClipPrinter\clip.exe' % USER_NAME
    add_to_startup(USER_NAME,file_path=dst)
    

def add_to_startup(USER_NAME,file_path):
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Run', 0, winreg.KEY_SET_VALUE);
    winreg.SetValueEx(key, 'ClipPrinter', 0, winreg.REG_SZ, file_path);
    

def main():
    copy_script()
    systray.start()
    timer = set_interval(getClip, 2)
    

if __name__ == "__main__":
    main()
