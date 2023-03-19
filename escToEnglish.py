#!/usr/bin/env python3

# pip install pynput
# pip install pywin32

from pynput.keyboard import Key, Listener, Controller
from os import system
import os
import signal
import ctypes
from ctypes import wintypes
import win32api
import win32process
import win32gui
import time

KEY = Key.esc
keyboard = Controller()

user32 = ctypes.WinDLL(name="user32")
imm32 = ctypes.WinDLL(name="imm32")

class GUITHREADINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("hwndActive", wintypes.HWND),
        ("hwndFocus", wintypes.HWND),
        ("hwndCapture", wintypes.HWND),
        ("hwndMenuOwner", wintypes.HWND),
        ("hwndMoveSize", wintypes.HWND),
        ("hwndCaret", wintypes.HWND),
        ("rcCaret", wintypes.RECT),

    ]

    def __str__(self):
        ret = "\n" + self.__repr__()
        start_format = "\n  {0:s}: "
        for field_name, _ in self. _fields_[:-1]:
            field_value = getattr(self, field_name)
            field_format = start_format + ("0d{1:016d}" if field_value else "{1:}")
            ret += field_format.format(field_name, field_value)
        rc_caret = getattr(self, self. _fields_[-1][0])
        ret += (start_format + "({1:d}, {2:d}, {3:d}, {4:d})").format(self. _fields_[-1][0], rc_caret.top, rc_caret.left, rc_caret.right, rc_caret.bottom)
        return ret


def handler(signum, frame):
    print("get signal")
    os._exit(-1)

def GetGUIThreadInfo(win32, tid, info):
    GetGUIThreadInfo_func = getattr(win32, "GetGUIThreadInfo")
    GetGUIThreadInfo_func.argtypes = [wintypes.DWORD, ctypes.POINTER(GUITHREADINFO)]
    GetGUIThreadInfo_func.restype = wintypes.BOOL
    info.cbSize = ctypes.sizeof(GUITHREADINFO)
    return GetGUIThreadInfo_func(tid, ctypes.byref(info))

def set_window_english(h_wnd, tmp):
    h_imc = imm32.ImmGetDefaultIMEWnd(h_wnd)
    WM_IME_CONTROL=643
    IMC_GETCONVERSIONMODE=0x01
    status = win32api.SendMessage(h_imc, WM_IME_CONTROL, IMC_GETCONVERSIONMODE, 0)
    # status != 0: means korean.
    # status == 0: means english.
    print(h_wnd, win32gui.GetWindowText(h_wnd), win32gui.GetClassName(h_wnd), status)
    if status != 0:
        # IMC_SETCONVERSIONMODE=2
        ret = win32api.SendMessage(h_imc, WM_IME_CONTROL, 2, 0) 
    
def change_to_english():
    h_wnd = win32gui.GetForegroundWindow()
    tid, pid = win32process.GetWindowThreadProcessId(h_wnd)
    info = GUITHREADINFO()
    if GetGUIThreadInfo(user32, tid, info):
        #print(info)
        #print(win32gui.GetWindowText(info.hwndFocus), win32gui.GetClassName(info.hwndFocus))
        set_window_english(info.hwndFocus, None)

def on_press(key):
    pass
        
def on_release(key):
    if key == KEY:
        change_to_english()
       
signal.signal(signal.SIGINT, handler)

listener = Listener(
        #on_press=on_press,
        on_release=on_release)

listener.start()
while True:
    #pass
    time.sleep(100)