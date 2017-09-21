"""
Clipboard Watcher with Threading using Pyperclip

Purpose:
    monitor the clipboard and perform changes or actions

Thanks to: 
    Pyperclip by Al Sweigart al@inventwithpython.com
    https://github.com/asweigart/pyperclip

Stackoverflow
    http://stackoverflow.com/a/14687465/4222833 
    http://stackoverflow.com/users/1156006/thorsten-kranz
    http://stackoverflow.com/users/1244729/kuanyui
 """

import time
import threading
import re
import pyperclip
import win32gui
import os
import webbrowser

### Todo make configurable
### Windows/Linux will be different
chrome = 'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe'

def clean_whitespaces(clipboard_content):
    ReSearchFilter = r'\n';
    ReSearch = r'^\s*(.*?)\s*$';
    tmp_clipboard_content = clipboard_content

    # Don't change content if it comes from these
    if re.search(r'\sExcel$', getActiveWindowTitle()):
        return False

    if re.search(r'Sublime', getActiveWindowTitle()):
        return False

    ### copied a Link
    # if re.search(r'Special skype window', getActiveWindowTitle()) and re.search(r'^\s*http', clipboard_content):
    #     clipboard_content = re.sub(ReSearch, r'\1', clipboard_content)
    #     os.system('"' + chrome + '" %s ' % clipboard_content)
    #     return True

    if 'TConversationForm' == getActiveWindowClass() and re.search(r'^\s*http', clipboard_content):
        clipboard_content = re.sub(ReSearch, r'\1', clipboard_content)
        os.system('"' + chrome + '" %s ' % clipboard_content)
        return True

    if not re.search(ReSearchFilter, clipboard_content) and re.search(ReSearch, clipboard_content):
        clipboard_content = re.sub(ReSearch, r'\1', clipboard_content)
        if tmp_clipboard_content == clipboard_content:
            return False
        setClipboard(clipboard_content)
        return True 
    return False

def print_to_stdout(clipboard_content):
    return
    clipboard_content = str((clipboard_content).encode('utf-8').strip())
    print "Window: '%s'" % (getActiveWindowTitle()).encode('utf-8').strip()
    print "Print: '%s'" % clipboard_content


def getActiveWindowTitle():
    try:
        return win32gui.GetWindowText (win32gui.GetForegroundWindow())
    except (RuntimeError, TypeError, NameError):
        return ''

def getActiveWindowClass():
    try:
        return win32gui.GetClassName (win32gui.GetForegroundWindow())
    except (RuntimeError, TypeError, NameError):
        return ''

class ClipboardWatcher(threading.Thread):
    def __init__(self, predicate, callback, pause=5.):
        super(ClipboardWatcher, self).__init__()
        self._predicate = predicate
        self._callback = callback
        self._pause = pause
        self._stopping = False

    def run(self):       
        recent_value = ""
        while not self._stopping:
            tmp_value = getClipboard()
            if tmp_value != "" and tmp_value != recent_value and not re.search(r'\sExcel$', getActiveWindowTitle()):
                recent_value = tmp_value
                if self._predicate(recent_value):
                    self._callback(recent_value)
                    recent_value = getClipboard()
            time.sleep(self._pause)

    def stop(self):
        self._stopping = True

def getClipboard():
    try:
        return pyperclip.paste()
    except (RuntimeError, TypeError, NameError):
        return ""

def setClipboard(clipboard_content):
    try:
        return pyperclip.copy(clipboard_content)
    except (RuntimeError, TypeError, NameError):
        return ""

def main():
    watcher = ClipboardWatcher(clean_whitespaces, 
                               print_to_stdout,
                               1.)
    watcher.start()
    while True:
        try:
            # print "Waiting for changed clipboard..."
            time.sleep(10)
        except KeyboardInterrupt:
            watcher.stop()
            break


if __name__ == "__main__":
    main()