# -*- coding: utf-8 -*-
"""
Functions in this module correspond more or less to the functions described 
in the HFSS Scripting Guide, Section "Reporter Editor Script Commands"

At last count there were 2 functions implemented out of 28.
"""
from __future__ import division, print_function, unicode_literals, absolute_import

from hycohanz.design import get_module


def export_to_file(design, reportname, filename):
    """
    From a data table or plot, generates text format, comma delimited, tab delimited, or .dat type output files.

    Parameters
    ----------
    design : pywin32 COMObject
        The HFSS design object upon which to operate.
    reportname : string
        The name of the report.
    filename : string
        Path and file name. The extension determines the type of file exported.
        - .txt : Post processor format file
        - .csv : Comma-delimited data file
        - .tab : Tab-separated file
        - .dat : Ansoft plot data file

    Returns
    -------
    None
    """
    module = get_module(design, "ReportSetup")
    module.ExportToFile(reportname, filename)


def get_all_report_names(design):
    """
    Gets the names of existing reports in a design.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design object upon which to operate.

    Returns
    -------
    reportnames = list of str
        A list with the names of all the reports in the Design
    """
    module = get_module(design, "ReportSetup")
    report_list = list(module.GetAllReportNames())
    return map(str, report_list)
