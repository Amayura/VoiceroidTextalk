# -*- coding: utf-8 -*-

import sys
import comtypes
from comtypes import CoCreateInstance
import comtypes.client
from comtypes.gen.UIAutomationClient import *

__uia = None
__root_element = None

def __init():
    global __uia, __root_element
    __uia = CoCreateInstance(CUIAutomation._reg_clsid_,
                             interface=IUIAutomation,
                             clsctx=comtypes.CLSCTX_INPROC_SERVER)
    __root_element = __uia.GetRootElement()

def get_window_element(title):
    win_element = __root_element.FindFirst(TreeScope_Children,
                                           __uia.CreatePropertyCondition(
                                           UIA_NamePropertyId, title))
    return win_element

def find_control(base_element, ctltype):
    condition = __uia.CreatePropertyCondition(UIA_ControlTypePropertyId, ctltype)
    ctl_elements = base_element.FindAll(TreeScope_Subtree, condition)
    return [ ctl_elements.GetElement(i) for i in range(ctl_elements.Length) ]

def find_child_control(element, ctltype):
    condition = __uia.CreatePropertyCondition(UIA_ControlTypePropertyId, ctltype)
    ctl_elements = base_element.FindAll(TreeScope_Subtree, condition)
    return [ ctl_elements.GetElement(i) for i in range(ctl_elements.Length) ]

def lookup_by_name(elements, name):
    for element in elements:
        if element.CurrentName == name:
            return element

    return None

def lookup_by_automationid(elements, id):
    for element in elements:
        if element.CurrentAutomationId == id:
            return element
    return None

def click_button(element):
    isClickable = element.GetCurrentPropertyValue(UIA_IsInvokePatternAvailablePropertyId)
    if isClickable == True:
        ptn = element.GetCurrentPattern(UIA_InvokePatternId)
        ptn.QueryInterface(IUIAutomationInvokePattern).Invoke()

def set_text(element, input_text):
    isValue = element.GetCurrentPropertyValue(UIA_IsValuePatternAvailablePropertyId)
    if isValue == True:
        ptn = element.GetCurrentPattern(UIA_ValuePatternId)
        ptn.QueryInterface(IUIAutomationValuePattern).SetValue(input_text)

def talk_voiceroid2(text):
    win = get_window_element('VOICEROID2')
    if not win: win = get_window_element('VOICEROID2*')
    if not win: sys.exit()

    btns = find_control(win, UIA_EditControlTypeId)
    set_text(btns[0], text)

    btns = find_control(win, UIA_ButtonControlTypeId)
    click_button(btns[7])

if __name__ == '__main__':
    if len(sys.argv) <= 1:
        sys.exit()
    
    __init()
    talk_voiceroid2(sys.argv[1])