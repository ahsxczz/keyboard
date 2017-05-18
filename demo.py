#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
from tkinter import Tk, Label, Entry, StringVar, Button
import pythoncom
import pyHook

PASSWORD = '230032'


def on_exit():
    if text.get() == PASSWORD:
        hm.UnhookKeyboard()
        root.withdraw()
        root.quit()
        sys.exit(1)
    else:
        text.set('')


def onKeyboardEvent(event):
    # password should be in 2, 3, 0
    return str(event.Key) in ('F5', '2', '3', '0', 'Back')


hm = pyHook.HookManager()
hm.KeyDown = onKeyboardEvent
hm.HookKeyboard()

root = Tk(className='Hook Keyboard')
width, height = 280, 150
root.geometry('%dx%d+400+200' % (width, height))
root.minsize(width, height)
root.maxsize(width, height)
root.protocol('WM_DELETE_WINDOW', on_exit)
Label(root, text='please enter password to exit').pack()
text = StringVar()
text.set('')
entry = Entry(root, show="*")
entry['textvariable'] = text
entry.pack()
entry.focus()
button = Button(root, text='Quit', command=on_exit).pack()
root.mainloop()  # 创建输入窗口，尚未实现无法关闭/关闭后循环弹出

pythoncom.PumpMessages()
