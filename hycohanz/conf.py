"""
This module is used as a global configuration file, where the global
COM objects are stored, and functions related to them are implemented.

All the implemented functions are decorators/wrappers that enable
calling each hycohanz function without their first COM object argument
(oDesign, oEditor, oProject, etc.) and, instead, using the global COM
objects defined here, that are updated by all the related functions internally.

Some COM objects are stored in lists in order to be able to completely clean
the hycohanz interface even though more than one project, design or editor
have been handled.
"""

import win32com.client

oDesktop = None
oAnsoftApp = None
oProjectList = []
oDesignList = []
oEditorList = []

## The following functions handle the internal storage of the COM objects
# When one of these functions is called from another hycohanz function, the
# new COM object becomes the currently handled internally

def update_oAnsoftApp(new_oAnsoftApp):
    oAnsoftApp = new_oAnsoftApp

def update_oDesktop(new_oDesktop):
    oDesktop = new_oDesktop

def update_oProject(new_oProject):
    if new_oProject not in oProjectList:
        oProjectList.append(new_oProject)
    else:
        oProjectList.append(oProjectList.pop(oProjectList.index(new_oProject)))

def update_oDesign(new_oDesign):
    if new_oDesign not in oDesignList:
        oDesignList.append(new_oDesign)
    else:
        oDesignList.append(oDesignList.pop(oDesignList.index(new_oDesign)))

def update_oEditor(new_oEditor):
    if new_oEditor not in oEditorList:
        oEditorList.append(new_oEditor)
    else:
        oEditorList.append(oEditorList.pop(oEditorList.index(new_oEditor)))

## Wrappers

def checkDefaultDesktop(func):
    """
    Decorator that makes the function use the global oDesktop object if
    no oDesktop win32 COM object is passed as an argument to the function
    """
    global oDesktop

    def wrapper(*args, **kwargs):
        if len(args)==0 or not isinstance(args[0], win32com.client.CDispatch):
            # Check if internal COM object already exists
            if not oDesktop:
                raise Exception("Internal oDesktop object has not been initialized yet")
            args = (oDesktop, *args)
        return func(*args, **kwargs)

    return wrapper

def checkDefaultProject(func):
    """
    Decorator that makes the function use the global oProject object if
    no oProject win32 COM object is passed as an argument to the function
    """
    global oProjectList

    def wrapper(*args, **kwargs):
        if len(args)==0 or not isinstance(args[0], win32com.client.CDispatch):
            # Check if internal COM object already exists
            if not oProjectList:
                raise Exception("Internal oProject object has not been initialized yet")
            args = (oProjectList[-1], *args)
        return func(*args, **kwargs)

    return wrapper

def checkDefaultDesign(func):
    """
    Decorator that makes the function use the global oDesign object if
    no oDesign win32 COM object is passed as an argument to the function
    """
    global oDesignList

    def wrapper(*args, **kwargs):
        if len(args)==0 or not isinstance(args[0], win32com.client.CDispatch):
            # Check if internal COM object already exists
            if not oDesignList:
                raise Exception("Internal oDesign object has not been initialized yet")
            args = (oDesignList[-1], *args)
        return func(*args, **kwargs)

    return wrapper

def checkDefaultEditor(func):
    """
    Decorator that makes the function use the global oEditor object if
    no oEditor win32 COM object is passed as an argument to the function
    """
    global oEditorList

    def wrapper(*args, **kwargs):
        if len(args)==0 or not isinstance(args[0], win32com.client.CDispatch):
            # Check if internal COM object already exists
            if not oEditorList:
                raise Exception("Internal oEditor object has not been initialized yet")
            args = (oEditorList[-1], *args)
        return func(*args, **kwargs)

    return wrapper
