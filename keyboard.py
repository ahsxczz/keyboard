# -*- coding: utf-8 -*- # 
from Tkinter import *
import pythoncom 
import pyHook#需要pyhook和pywin32

def onKeyboardEvent(event):
	if str(event.Key) in ('Escape','Subtract','Add','Return','Oem_Minus','Oem_Plus'):
		print ("Key:", event.Key) 
		print ("---")
  # 同鼠标事件监听函数的返回值   
		return True
	else:
		return True

hm = pyHook.HookManager()  
hm.KeyDown = onKeyboardEvent 
hm.HookKeyboard()
pythoncom.PumpMessages()