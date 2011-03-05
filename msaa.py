#!/usr/bin/env python
# coding:utf-8

import sys, os, re, time
import platform
import ctypes, ctypes.wintypes
import comtypes, comtypes.automation, comtypes.client

comtypes.client.GetModule('oleacc.dll')

AccRoleNameMap = {
    1: u'TitleBar',
    2: u'MenuBar',
    3: u'ScrollBar',
    4: u'Grip',
    5: u'Sound',
    6: u'Cursor',
    7: u'Caret',
    8: u'Alert',
    9: u'Window',
    10: u'Client',
    11: u'PopupMenu',
    12: u'MenuItem',
    13: u'Tooltip',
    14: u'Application',
    15: u'Document',
    16: u'Pane',
    17: u'Chart',
    18: u'Dialog',
    19: u'Border',
    20: u'Grouping',
    21: u'Separator',
    22: u'ToolBar',
    23: u'StatusBar',
    24: u'Table',
    25: u'ColumnHeader',
    26: u'RowHeader',
    27: u'Column',
    28: u'Row',
    29: u'Cell',
    30: u'Link',
    31: u'HelpBalloon',
    32: u'Character',
    33: u'List',
    34: u'ListItem',
    35: u'Outline',
    36: u'OutlineItem',
    37: u'PageTab',
    38: u'PropertyPage',
    39: u'Indicator',
    40: u'Graphic',
    41: u'Text',
    42: u'EditableText',
    43: u'PushButton',
    44: u'CheckBox',
    45: u'RadioButton',
    46: u'ComboBox',
    47: u'DropDown',
    48: u'ProgressBar',
    49: u'Dial',
    50: u'HotKeyField',
    51: u'Slider',
    52: u'SpinBox',
    53: u'Diagram',
    54: u'Animation',
    55: u'Equation',
    56: u'DropDownButton',
    57: u'MenuButton',
    58: u'GridDropDownButton',
    59: u'WhiteSpace',
    60: u'PageTabList',
    61: u'Clock',
    62: u'SplitButton',
    63: u'IPAddress',
}

class Element(object):
    '''Accessible Element'''
    def __init__(self, IAccessible, iObjectId):
        if not isinstance(IAccessible, comtypes.gen.Accessibility.IAccessible):
            raise TypeError(u'Element(IAccessible,iObjectId) first argument type must be IAccessible')
        if not isinstance(iObjectId, int):
            raise TypeError(u'Element(IAccessible,iObjectId) second argument type must be int')
        self.IAccessible = IAccessible
        self.iObjectId = iObjectId

    def accChildCount(self):
        if self.iObjectId == 0:
            return self.IAccessible.accChildCount
        else:
            return 0

    def accRole(self):
        objChildId = comtypes.automation.VARIANT()
        objChildId.vt = comtypes.automation.VT_I4
        objChildId.value = self.iObjectId
        objRole = comtypes.automation.VARIANT()
        objRole.vt = comtypes.automation.VT_BSTR
        self.IAccessible._IAccessible__com__get_accRole(objChildId, objRole)
        return objRole.value

    def accName(self):
        objChildId = comtypes.automation.VARIANT()
        objChildId.vt = comtypes.automation.VT_I4
        objChildId.value = self.iObjectId
        objName = comtypes.automation.BSTR()
        self.IAccessible._IAccessible__com__get_accName(objChildId, ctypes.byref(objName))
        return objName.value

    def accLocation(self):
        '''return (left, top, width, height)'''
        objChildId = comtypes.automation.VARIANT()
        objChildId.vt = comtypes.automation.VT_I4
        objChildId.value = self.iObjectId
        objL, objT, objW, objH = ctypes.c_long(), ctypes.c_long(), ctypes.c_long(), ctypes.c_long()
        self.IAccessible._IAccessible__com_accLocation(ctypes.byref(objL), ctypes.byref(objT), ctypes.byref(objW), ctypes.byref(objH), objChildId)
        return (objL.value, objT.value, objW.value, objH.value)

    def accDefaultAction(self):
        return self.IAccessible.accDefaultAction()

    def accDescription(self):
        return self.IAccessible.accDescription()

    def accHelp(self):
        return self.IAccessible.accHelp()

    def accHelpTopic(self):
        return self.IAccessible.accHelpTopic()

    def accKeyboardShortcut(self):
        return self.IAccessible.accKeyboardShortcut()

    def accNavigate(self):
        return self.IAccessible.accNavigate()

    def accParent(self):
        return self.IAccessible.accParent()

    def accState(self):
        return self.IAccessible.accState()

    def accValue(self):
        return self.IAccessible.accValue()

    def accDoDefaultAction(self):
        return self.IAccessible.accDoDefaultAction()

    def accFocus(self):
        if self.iObjectId:
            return self.IAccessible.accFocus(self.iObjectId)
        else:
            return self.IAccessible.accFocus()

    def accSelect(self, iSelection):
        '''
SELFLAG_TAKEFOCUS       1
SELFLAG_TAKESELECTION   2
SELFLAG_TAKESELECTION   2
SELFLAG_EXTENDSELECTION 4
SELFLAG_ADDSELECTION    8
SELFLAG_REMOVESELECTION 16
'''
        if self.iObjectId:
            return self.IAccessible.accSelect(iSelection, self.iObjectId)
        else:
            return self.IAccessible.accSelect(iSelection)

    def accRoleName(self):
        try:
            iRole = self.accRole()
            return AccRoleNameMap.get(iRole)
        except:
            return None

    def __iter__(self):
        if self.iObjectId > 0:
            raise StopIteration()
        objAccChildArray = (comtypes.automation.VARIANT * self.IAccessible.accChildCount)()
        objAccChildCount = ctypes.c_long()
        ctypes.oledll.oleacc.AccessibleChildren(self.IAccessible, 0, self.IAccessible.accChildCount, objAccChildArray, ctypes.byref(objAccChildCount))
        for i in xrange(objAccChildCount.value):
            objAccChild = objAccChildArray[i]
            if objAccChild.vt == comtypes.automation.VT_DISPATCH:
                yield Element(objAccChild.value.QueryInterface(comtypes.gen.Accessibility.IAccessible), 0)
            else: #if objAccChild.vt == comtypes.automation.VT_I4:
                yield Element(self.IAccessible, i+1)

    def __str__(self):
        iRole = self.accRole()
        return '[%s(0x%X)|%r|ChildCount:%d]' % (AccRoleNameMap.get(iRole, 'Unkown'), iRole, self.accName(), self.IAccessible.accChildCount)

    def finditer(self, strRoleName, **kwargs):
        lstStack = [self]
        while lstStack:
            objElement = lstStack.pop()
            if objElement.IAccessible.accChildCount > 0:
                lstStack += reversed(list(objElement))
            if objElement == self:
                continue
            if strRoleName:
                iRole = objElement.accRole()
                if AccRoleNameMap.get(iRole) != strRoleName:
                    continue
            bMatched = True
            for strProperty in kwargs:
                try:
                    attr = getattr(objElement, 'acc'+strProperty)
                except AttributeError:
                    continue
                try:
                    value = attr()
                except:
                    value = None
                if type(kwargs[strProperty]) == type(lambda x:True):
                    bMatched = kwargs[strProperty]
                    if not bMatched:
                        break
                else:
                    if value != kwargs[strProperty]:
                        bMatched = False
                        break
            if bMatched:
                yield objElement

    def find(self, strRoleName, **kwargs):
        try:
            return self.finditer(strRoleName, **kwargs).next()
        except StopIteration:
            return None
    def findall(self, strRoleName, **kwargs):
        return list(self.finditer(strRoleName, **kwargs))


def ElementFromWindow(strWindowClassName, strWindowTitle):
    if strWindowClassName is not None and not isinstance(strWindowClassName, basestring):
        raise TypeError(u'strWindowClassName must be a string or None')
    if strWindowTitle is not None and not isinstance(strWindowTitle, basestring):
        raise TypeError(u'strWindowTitle must be a string or None')
    strWindowClassName = unicode(strWindowClassName) if isinstance(strWindowClassName, str) else None
    strWindowTitle = unicode(strWindowTitle) if isinstance(strWindowTitle, str) else None
    iHwnd = ctypes.windll.user32.FindWindowW(strWindowClassName, strWindowTitle)
    if iHwnd == 0:
        raise RuntimeError(u'Cannot FindWindowW(%r, %r)' % (strWindowClassName, strWindowTitle))
    return ElementFromHWND(iHwnd)

def ElementFromHWND(iHwnd):
    IAccessible = ctypes.POINTER(comtypes.gen.Accessibility.IAccessible)()
    ctypes.oledll.oleacc.AccessibleObjectFromWindow(iHwnd, 0, ctypes.byref(comtypes.gen.Accessibility.IAccessible._iid_), ctypes.byref(IAccessible))
    return Element(IAccessible, 0)

def ElementFromPoint(x, y):
    objPoint = ctypes.wintypes.POINT()
    objPoint.x = x
    objPoint.y = y
    IAccessible = ctypes.POINTER(comtypes.gen.Accessibility.IAccessible)()
    objChildId = comtypes.automation.VARIANT()
    ctypes.oledll.oleacc.AccessibleObjectFromPoint(objPoint, ctypes.byref(IAccessible), ctypes.byref(objChildId))
    return Element(IAccessible, objChildId.value)


def ElementToXML(objElement):
    objRoot = lxml.etree.fromstring('<Accessible />')
    lstQueue = [(objElement, objRoot)]
    while lstQueue:
        objElement, objTree = lstQueue.pop(0)
        objRoleName = objElement.accRoleName()
        objName = objElement.accName()
        strRoleName = objRoleName or 'Unkown'
        strName = cgi.escape(unicode(objName)) if objName else ''
        #strLocation = unicode(objLocation).strip('()').replace(' ', '')
        objSubTree = lxml.etree.SubElement(objTree, strRoleName)
        try:
            objSubTree.attrib['Name'] = strName
        except:
            objSubTree.attrib['Name'] = strName.encode('unicode-escape')
        #objSubTree.attrib['Location'] = strLocation
        if objElement.IAccessible.accChildCount > 0:
            for objElementChild in objElement:
                lstQueue.append((objElementChild, objSubTree))
    return lxml.etree.tostring(objRoot, pretty_print=1)

def Find(objHandle, strRoleName, **kwargs):
    if isinstance(objHandle, Element):
        objElement = objHandle
    elif objHandle in (0, None):
        objElement = ElementFromHWND(ctypes.windll.user32.GetDesktopWindow())
    elif isinstance(objHandle, (int, long)):
        objElement = ElementFromHWND(objHandle)
    elif isinstance(objHandle, basestring):
        iHwnd = ctypes.windll.user32.FindWindowW(objHandle, None) or ctypes.windll.user32.FindWindowW(None, objHandle)
        assert iHwnd > 0
        objElement = ElementFromHWND(iHwnd)
    else:
        raise TypeError(u'AccFind objHandle must be a int/str/unicode/AccObject, not %r' % objHandle)
    return objElement.find(strRoleName, **kwargs)

