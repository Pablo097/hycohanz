# -*- coding: utf-8 -*-
"""
Functions in this module correspond more or less to the functions described
in the HFSS Scripting Guide, Section "Desktop Object Script Commands".

At last count there were 11 functions implemented out of 31.
"""
from __future__ import division, print_function, unicode_literals, absolute_import

import hycohanz.conf as conf
from hycohanz.project import get_project_name

@conf.checkDefaultDesktop
def quit_application(oDesktop):
    """
    Exit HFSS.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        HFSS Desktop object.

    Returns
    -------
    None

    Examples
    --------
    >>> import Hyphasis as hfss
    >>> [oAnsoftApp, oDesktop] = hfss.setup_interface()
    >>> hfss.quit_application(oDesktop)

    """
    oDesktop.QuitApplication()

@conf.checkDefaultDesktop
def new_project(oDesktop):
    """
    Create a new project.  The new project becomes the active project.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        HFSS Desktop object in which to create the new project.

    Returns
    -------
    oProject : pywin32 COMObject
        The created HFSS project.

    """
    conf.oProject = oDesktop.NewProject()

    return conf.oProject

@conf.checkDefaultDesktop
def open_project(oDesktop, filename):
    """
    Open an HFSS project.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object to which this function is applied.
    filename : str
        The name of the file to open.

    Returns
    -------
    oProject : pywin32 COMObject
        An handle to the opened project.

    """
    conf.oProject = oDesktop.OpenProject(filename)
    return conf.oProject

@conf.checkDefaultDesktop
def close_project_byname(oDesktop, projectname):
    """
    Close the specified project using the given HFSS project object.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object upon which to operate.
    projectname : string
        The name of the HFSS project to close.

    Returns
    -------
    None

    """
    oDesktop.CloseProject(projectname)

@conf.checkDefaultDesktop
def get_active_project(oDesktop):
    """
    Get a handle for the active project.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object upon which to operate.

    Returns
    -------
    oProject : pywin32 COMObject
        The HFSS project object upon which to operate.

    """
    conf.oProject = oDesktop.GetActiveProject()

    return conf.oProject

@conf.checkDefaultDesktop
def set_active_project(oDesktop, projectname):
    """
    Set the specific project as the active one.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object upon which to operate.
    projectname : string
        The name of the HFSS project to set as active.

    Returns
    -------
    oProject : pywin32 COMObject
        The HFSS project object upon which to operate.

    """
    conf.oProject = oDesktop.SetActiveProject(projectname)

    return conf.oProject

@conf.checkDefaultDesktop
def close_project_byhandle(oDesktop, oProject):
    """
    Close the specified project using the given HFSS project object.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object upon which to operate.
    oProject : pywin32 COMObject
        The HFSS project object upon which to operate.

    Returns
    -------
    None

    """
    oDesktop.CloseProject(get_project_name(oProject))

@conf.checkDefaultDesktop
def close_current_project(oDesktop):
    """
    Close the active project.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object upon which to operate.

    Returns
    -------
    None

    """
    oProject = get_active_project(oDesktop)
    projectname = get_project_name(oProject)
    oDesktop.CloseProject(projectname)

@conf.checkDefaultDesktop
def get_projects(oDesktop):
    """
    Get the list of open projects.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object upon which to operate.

    Returns
    -------
    oProject : list of pywin32 COMObjects
        The HFSS desktop object upon which to operate.

    """
    return oDesktop.GetProjects()

@conf.checkDefaultDesktop
def close_all_projects(oDesktop):
    """
    Close all open projects.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object upon which to operate.

    Returns
    -------
    None

    """
    objlist = get_projects(oDesktop)

    for item in objlist:
        close_project_byhandle(oDesktop, item)

@conf.checkDefaultDesktop
def close_all_projects_except_current(oDesktop):
    """
    Close all open projects except the active project.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object upon which to operate.

    Returns
    -------
    None

    """
    currproj = get_project_name(get_active_project(oDesktop))
    projlist = get_projects(oDesktop)

    for item in projlist:
        if get_project_name(item) != currproj:
            close_project_byhandle(oDesktop, item)

@conf.checkDefaultDesktop
def save_as_project(oDesktop, filename,overwrite=True):
    """
    Save As an HFSS project.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        The HFSS desktop object to which this function is applied.
    filename : str
        The name of the file to save, can include path. Works well with python os paths (import os in your main project)

    Returns
    -------
    oProject : pywin32 COMObject
        An handle to the saved project.

    """
    oProject = get_active_project(oDesktop)
    return oProject.SaveAs(filename,overwrite)
