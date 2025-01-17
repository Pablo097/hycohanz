# -*- coding: utf-8 -*-
"""
Functions in this module correspond more or less to the functions described
in the HFSS Scripting Guide, Section "Ansoft Application Object Script
Commands".

At last count there were 1 functions implemented out of 15.
"""
from __future__ import division, print_function, unicode_literals, absolute_import

import win32com.client

import hycohanz.conf as conf

def setup_interface():
    """
    Set up the COM interface to the running HFSS process.

    Returns
    -------
    oAnsoftApp : pywin32 COMObject
        Handle to the HFSS application interface
    oDesktop : pywin32 COMObject
        Handle to the HFSS desktop interface

    Examples
    --------
    >>> import hycohanz as hfss
    >>> [oAnsoftApp, oDesktop] = hfss.setup_interface()

    """
    # I'm still looking for a better way to do this.  This attaches to an
    # existing HFSS process instead of creating a new one.  I would highly
    # prefer that a new process is created.  Apparently
    # win32com.client.DispatchEx() doesn't work here either.
    conf.oAnsoftApp = win32com.client.Dispatch('AnsoftHfss.HfssScriptInterface')

    conf.oDesktop = conf.oAnsoftApp.GetAppDesktop()

    return [conf.oAnsoftApp, conf.oDesktop]

def clean_interface():
    """
    This function should be called at the end of each script that has executed
    the setup_interface() function.
    """
    del conf.oDesktop
    del conf.oAnsoftApp
    # del conf.oProject
    # del conf.oDesign
    # del conf.oEditor
    for oProject in conf.oProjectList:
        del oProject
    for oDesign in conf.oDesignList:
        del oDesign
    for oEditor in conf.oEditorList:
        del oEditor
