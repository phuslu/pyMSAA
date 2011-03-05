#!/usr/bin/env python
# coding:utf-8

import sys, os, re
import comtypes.client

AutoItX = None

def autoit():
    global AutoItX
    if AutoItX is None:
        strDllName = 'AutoItX3.dll'
        if platform.architecture()[0] == '64bit':
            strDllName = 'AutoItX3_x64.dll'
        ctypes.WinDLL(os.path.join(os.path.dirname(__file__), strDllName)).DllRegisterServer()
        AutoItX = comtypes.client.CreateObject('AutoItX3.Control')
    return AutoItX



