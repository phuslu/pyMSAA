#!/usr/bin/env python
# coding:utf-8

import sys, os, re, logging, time
import ctypes, ctypes.wintypes
import msaa, comtypes.client

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

def GetCursorPos():
    objPoint = ctypes.wintypes.POINT()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(objPoint))
    return objPoint.x, objPoint.y

def GetCurrentElementInfo(objElement):
    dictInfo = {}
    dictInfo['ChildId'] = objElement.iObjectId
    lstAttributeNameList = [
                            'accRoleName',
                            'accRole',
                            'accName',
                            'accValue',
                            'accState',
                            'accLocation',
                            'accDescription',
                            'accKeyboardShortcut',
                            'accDefaultAction',
                            'accHelp',
                            'accHelpTopic',
                            'accChildCount'
                            ]
    for attr in lstAttributeNameList:
        try:
            dictInfo[attr] = getattr(objElement, attr)()
        except:
            dictInfo[attr] = None
    return '\n'.join('%s:\t%r' % (attr, dictInfo[attr]) for attr in lstAttributeNameList)

def main():
    x_old, y_old = GetCursorPos()
    while 1:
        x, y = GetCursorPos()
        if (x, y) != (x_old, y_old):
            x_old, y_old = x, y
            objElement = msaa.point(x, y)
            os.system('cls')
            print GetCurrentElementInfo(objElement)
        time.sleep(0.5)

if __name__ == '__main__':
    main()


