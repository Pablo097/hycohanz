"""
This module is used as a global configuration file, where the global
COM objects are stored, and functions related to them are implemented.

All the implemented functions are decorators/wrappers that enable
calling each hycohanz function without their first COM object argument
(oDesign, oEditor, oProject, etc.) and, instead, using the global COM
objects defined here, that are updated by all the related functions internally.
"""

import win32com.client

oDesktop = None
oAnsoftApp = None
oProject = None
oDesign = None
oEditor = None

def checkDefaultDesktop(func):
    """
    Decorator that makes the function use the global oDesktop object if
    no oDesktop win32 COM object is passed as an argument to the function
    """
    global oDesktop
    def wrapper(*args, **kwargs):
        if len(args)==0 or not isinstance(args[0], win32com.client.CDispatch):
            args = (oDesktop, *args)
        return func(*args, **kwargs)

    return wrapper

def checkDefaultProject(func):
    """
    Decorator that makes the function use the global oProject object if
    no oProject win32 COM object is passed as an argument to the function
    """
    global oProject
    def wrapper(*args, **kwargs):
        if len(args)==0 or not isinstance(args[0], win32com.client.CDispatch):
            args = (oProject, *args)
        return func(*args, **kwargs)

    return wrapper

def checkDefaultDesign(func):
    """
    Decorator that makes the function use the global oDesign object if
    no oDesign win32 COM object is passed as an argument to the function
    """
    global oDesign
    def wrapper(*args, **kwargs):
        if len(args)==0 or not isinstance(args[0], win32com.client.CDispatch):
            args = (oDesign, *args)
        return func(*args, **kwargs)

    return wrapper

def checkDefaultEditor(func):
    """
    Decorator that makes the function use the global oEditor object if
    no oEditor win32 COM object is passed as an argument to the function
    """
    global oEditor
    def wrapper(*args, **kwargs):
        if len(args)==0 or not isinstance(args[0], win32com.client.CDispatch):
            args = (oEditor, *args)
        return func(*args, **kwargs)

    return wrapper
