#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import unicode_literals
import sys
import os
from Tkinter import Tk, Label, Entry, StringVar, Button
import pythoncom
import pyHook
#import psutil
import ctypes
import win32gui
import win32api
import win32con
import win32process
import win32com.client
import time

FS_CODING = sys.getfilesystemencoding()
EnumWindows = ctypes.windll.user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int,
                                     ctypes.POINTER(ctypes.c_int))
GetWindowText = ctypes.windll.user32.GetWindowTextW
GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
IsWindowVisible = ctypes.windll.user32.IsWindowVisible
SetWindowText = ctypes.windll.user32.SetWindowTextA
SetCursorPos = ctypes.windll.user32.SetCursorPos

keystr = 1
PASSWORD = '330033'

def on_exit():
    if text.get() == PASSWORD:
        hm.UnhookKeyboard()
        root.withdraw()#隐藏窗口
        root.quit()
        sys.exit(1)
    else:
        text.set('')
		
def on_click():
	global keystr
	update_hook()
	if text.get() == PASSWORD:
		if keystr==0:
			keystr=1
		else:
			keystr=0
		
def get_child(parent=0, child=None, cls_name=None, window_name=None):
    return win32gui.FindWindowEx(parent, child, cls_name, window_name)


def get_text(hwnd, safe=True, safe_text=''):
    try:
        buf_size = 1 + win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH,
                                            0, 0)
        buf = win32gui.PyMakeBuffer(buf_size)
        win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, buf_size, buf)
        return buf[:buf_size - 1]
    except:
        if safe:
            return safe_text

        raise


def compare_text(src, dst, rule='=', re_rule=None):
    if rule == 'startswith':
        return src.startswith(dst)
    elif rule == 'endswith':
        return src.endswith(dst)
    elif rule == 're':
        return re_rule.match(src)
    else:
        return src == dst


def check_pid(hwnd, pid):
    if pid in win32process.GetWindowThreadProcessId(hwnd):
        bring_foreground(hwnd)


def show_window_by_pid(pid):
    return win32gui.EnumWindows(check_pid, pid)


def bring_foreground(hwnd):
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('%')
    win32gui.SetForegroundWindow(hwnd)


def show_window(hwnd, cmd=None):
    if cmd is None:
        cmd = win32con.SW_SHOW

    win32gui.ShowWindow(hwnd, cmd)


def get_handler_text(hwnd):
    length = GetWindowTextLength(hwnd)
    buff = ctypes.create_unicode_buffer(length + 1)
    GetWindowText(hwnd, buff, length + 1)
    return buff.value


def get_all_handlers(visible=True):
    handlers = []

    def foreach_window(hwnd, l_param):
        if not visible or IsWindowVisible(hwnd):
            handlers.append((hwnd, get_handler_text(hwnd)))

        return True

    EnumWindows(EnumWindowsProc(foreach_window), 0)
    return handlers


def find_handler(dst, rule='=', re_rule=None, debug=False, visible=True):
    for hwnd, text in get_all_handlers(visible):
        if compare_text(text, dst, rule, re_rule):
            return hwnd

    return None


def is_visible(hwnd):
    return win32gui.IsWindowVisible(hwnd)


def press_button(hwnd):
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, 0)
    win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, 0, 0)


def close_window(hwnd, force=False):
    if force or is_visible(hwnd):
        win32gui.SendMessage(hwnd, win32con.WM_CLOSE, 0, 0)


def get_foreground_window():
    return win32gui.GetForegroundWindow()


def get_window_text(hwnd, safe_text=''):
    return win32gui.GetWindowText(hwnd) or safe_text


def get_class_name(hwnd, safe=True):
    try:
        return win32gui.GetClassName(hwnd)
    except:
        if safe:
            return ''

        raise


def get_window_pos_and_size(hwnd):
    rect = win32gui.GetWindowRect(hwnd)
    x, y = rect[0], rect[1]
    return x, y, rect[2] - x, rect[3] - y


def move_to(hwnd, x, y, w, h, repaint=True):
    win32gui.MoveWindow(hwnd, x, y, w, h, repaint)


def set_cursor_pos(x, y):
    SetCursorPos(int(x), int(y))


def set_text(hwnd, text):
    SetWindowText(hwnd, ctypes.c_char_p(text))


def send_chars(hwnd, text, enter=False):
    [win32gui.SendMessage(hwnd, win32con.WM_CHAR, ord(c), 0) for c in text]
    if enter:
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)


def send_text(hwnd, text, enter=False):
    if isinstance(text, unicode):
        text = text.encode(FS_CODING)

    win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, 0, text)
    if enter:
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)


def enable_handler(hwnd, enable=True):
    return win32gui.EnableWindow(hwnd, enable)


def click_button(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, 0, 0)
    win32api.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, 0)

def get_all_handlers_hwnd(visible=True):
    handlers_hwnd = []
	
    def foreach_window(hwnd, l_param):
        if not visible or IsWindowVisible(hwnd):
            handlers_hwnd.append((hwnd))
        return True

    EnumWindows(EnumWindowsProc(foreach_window), 0)
    return handlers_hwnd

def get_all_handlers_id(visible=True):
    handlers_id = []
	
    def foreach_window(hwnd, l_param):
        if not visible or IsWindowVisible(hwnd):
            handlers_id.append((get_handler_text(hwnd)))
        return True

    EnumWindows(EnumWindowsProc(foreach_window), 0)
    return handlers_id

def is_enabled(hwnd):
    return win32gui.IsWindowEnabled(hwnd)

def SetAsForegroundWindow(hwnd):
	win32gui.SetForegroundWindow(hwnd)

def get_processes(name=None, attrs=['pid', 'name']):
    processes = []
    for proc in psutil.process_iter():
        try:
            pinfo = proc.as_dict(attrs=attrs)
        except psutil.NoSuchProcess:
            pass
        else:
            if name:
                if pinfo['name'] == name:
                    processes.append(pinfo)

                continue

            processes.append(pinfo)

    return processes	
	
def onKeyboardEvent(event):
    z1=get_all_handlers_id()
    z2=get_all_handlers_hwnd()
    for i in range(len(z1)):
        if z1[i].startswith('APT-9401SL 5.0-1b'):
            a=z2[i]
            break
        else:
            a=None
    # password should be in Y, 3, 0 , + , - ,=
    #return str(event.Key) in ('F5', 'Y', '3', '0', 'Back','Numpad3','Numpad0','Escape','Subtract','Add','Return','Oem_Minus','Oem_Plus','Up','Down','Left','Right')
    if str(event.Key) in ('1','2','9','Numpad1','Numpad2','Numpad9','Decimal','Oem_Period'):
        if a==get_foreground_window():
            return False
        elif find_handler('系列番号輸入'):
            return True
        elif find_handler('分組輸入系列番号'):
            return True
    else:
        return True
	
def update_hook():
	x=get_all_handlers_id()
	y=get_all_handlers_hwnd()
	for i in range(100):
		if x[i].startswith('APT-9401SL 5.0-1b'):
			a=y[i]
			break
	b1=(get_child(parent=a))
	for i in range(6):
		b1=(get_child(parent=a,child=b1))
	c1=(get_child(parent=b1))
	d1=(get_child(parent=c1))
	d2=(get_child(parent=c1,child=d1))
	e1=(get_child(parent=d2))
	f1=(get_child(parent=e1))
	g1=(get_child(parent=f1))
	for i in range(26):
		g1=(get_child(parent=f1,child=g1))#27
	g2=(get_child(parent=f1,child=g1))#28
	g3=(get_child(parent=f1,child=g2))
	for i in range(2):
		g3=(get_child(parent=f1,child=g3))#31
	h1=(get_child(parent=g3))
	h2=(get_child(parent=g3,child=h1))#31-2
	g4=(get_child(parent=f1,child=g3))#32
	g5=(get_child(parent=f1,child=g4))#33
	g6=(get_child(parent=f1,child=g5))#34
	g7=(get_child(parent=f1,child=g6))#35
	if (keystr):
		enable_handler(g1,False)
		enable_handler(g2,False)
		enable_handler(h2,False)
		enable_handler(g4,False)
		enable_handler(g5,False)
		enable_handler(g6,False)
		enable_handler(g7,False)
	else:
		enable_handler(g1,True)
		enable_handler(g2,True)
		enable_handler(h2,True)
		enable_handler(g4,True)
		enable_handler(g5,True)
		enable_handler(g6,True)
		enable_handler(g7,True)
	root.after(1000,update_hook)
	
def close_taskmanager():
	#if get_processes('VsTskMgr.exe'):
	if find_handler('Windows Task Manager'):
		os.system('taskkill /f /im VsTskMgr.exe')
	root.after(1000,close_taskmanager)

#def check_exsit(process_name):
    #WMI = win32com.client.GetObject('winmgmts:')
    #processCodeCov = WMI.ExecQuery('select * from Win32_Process where Name="%s"' % process_name)
    #if len(processCodeCov) > 0:
		#os.system('taskkill /f /im VsTskMgr.exe')

def check_update():
	check_exsit('VsTskMgr.exe')
	root.after(1000,check_update)
	
if __name__ == '__main__':	
	hm = pyHook.HookManager()
	hm.KeyDown = onKeyboardEvent
	hm.HookKeyboard()

	
	root = Tk(className='Hook Keyboard')
	width, height = 280, 150
	root.geometry('%dx%d+400+200' % (width, height))
	root.minsize(width, height)
	root.maxsize(width, height)
	root.protocol('WM_DELETE_WINDOW', on_exit)#on_exit点击关闭触发
	Label(root, text='please enter password to exit').pack()
	text = StringVar()
	text.set('')
	entry = Entry(root, show="*")
	entry['textvariable'] = text
	entry.pack()
	entry.focus()
	button2 = Button(root, text='Change', command=on_click).pack()
	button = Button(root, text='Quit', command=on_exit).pack()
	#close_taskmanager()
	#check_update()
	root.mainloop()  #消息循环
	
	pythoncom.PumpMessages()