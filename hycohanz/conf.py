import win32com.client

oDesktop = None
oAnsoftApp = None
oProject = None
oDesign = None
oEditor = None

def delete_all_COM_objects():
    global oDesktop
    del oDesktop
    global oAnsoftApp
    del oAnsoftApp
    global oProject
    del oProject
    global oDesign
    del oDesign
    global oEditor
    del oEditor

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
