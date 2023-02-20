# -*- coding: utf-8 -*-
"""
Functions in this module correspond more or less to the functions described
in the HFSS Scripting Guide, Section "Reporter Editor Script Commands"

At last count there were 5 functions implemented out of 28.
"""
from __future__ import division, print_function, unicode_literals, absolute_import

import hycohanz.conf as conf
from hycohanz.design import get_module
from hycohanz.analysis_setup import get_setups, get_sweeps

@conf.checkDefaultDesign
def create_report(oDesign,
                  report_name = "S Parameter Plot 1",
                  report_type = "Modal Solution Data",
                  display_type = "Rectangular Plot",
                  setup_name = "Setup1",
                  sweep_name = "LastAdaptive",
                  context_array = ["Domain:=", "Sweep"],
                  families_array = [],
                  report_data_array = ["X Component:=", "Freq", "Y Component:=", ["dB(S(P1,P1))"]]):
    """
    Creates a new report with a single trace and adds it to the Results branch in the project tree.

    oDesign : pywin32 COMObject
        The HFSS oDesign object upon which to operate.
    ReportName : string
        The name of the report.
    ReportType : string
        The type of report. Possible values are:
        "Modal Solution Data" - Only for Driven Modal solution-type problems with ports.
        "Terminal S Parameters" - Only for Driven Terminal solution-type problems with ports.
        "Eigenmode Parameters" - Only for Eigenmode solution-type problems.
        "Fields"
        "Far Fields" - Only for problems with radiation or PML boundaries.
        "Near Fields" - Only for problems with radiation or PML boundaries.
        “Emission Test”
    DisplayType : string
        If ReportType is "Modal S Parameters", "Terminal S Parameters", or "Eigenmode Parameters",
        then set to one of the following:
            "Rectangular Plot", "Polar Plot", "Radiation Pattern", "Smith Chart", "Data Table",
            "3D Rectangular Plot", or "3D Polar Plot".
        If ReportType is "Fields", then set to one of the following:
            "Rectangular Plot", "Polar Plot", "Radiation Pattern", "Data Table", or "3D Rectangular Plot".
        If ReportType is "Far Fields" or "Near Fields", then set to one of the following:
            "Rectangular Plot", "Radiation Pattern", "Data Table", "3D Rectangular Plot", or "3D Polar Plot"
        If ReportType is “Emission Test”, then set to one of the following:
            “Rectangular Plot” or “Data Table”
    SetupName : str
        Name of HFSS setup to use, for example "Setup1"
    SweepName : str
        Name of HFSS sweep to use, for example "LastAdaptive"
    ContextArray : list of strings
        Context for which the expression is being evaluated. This can be an empty string if there is no context.
        Must be passed in as a pair, for example
            ["Domain:=", "Sweep"]
            ["Domain:=", "Time"]
            ["Context:=", "Infinite Sphere"]
    FamiliesArray : list of strings
        Contains sweep definitions for the report.
        Must be passed in as a pair, for example
            ["VariableName:=", "Value"]
            ["Theta:=", "All")
    ReportDataArray : list of strings
        This array contains the report quantity and X, Y, and (Z) axis definitions.
        ["X Component:=", VariableName, "Y Component:=", VariableName]

    Returns
    -------
    None
    """

    #  TODO: Is there a way to improve the way these lists are input? (contextarray, familiesarray, reportdataarray)
    #  TODO: Provide some error checking on reporttype and displaytype

    check_setup(oDesign, setup_name)
    if sweep_name != "LastAdaptive":
        check_sweep(oDesign, setup_name, sweep_name)
    solution_name = setup_name + " : " + sweep_name

    oModule = get_module(oDesign, "ReportSetup")
    oModule.CreateReport(report_name,
                        report_type,
                        display_type,
                        solution_name,
                        context_array,
                        families_array,
                        report_data_array,
                        [])

@conf.checkDefaultDesign
def export_to_file(oDesign, report_name, filename):
    """
    From a data table or plot, generates text format, comma delimited,
    tab delimited, or .dat type output files.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS oDesign object upon which to operate.
    reportName : string
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
    oModule = get_module(oDesign, "ReportSetup")
    oModule.ExportToFile(report_name, filename)

@conf.checkDefaultDesign
def get_all_report_names(oDesign):
    """
    Gets the names of existing reports in a design.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS oDesign object upon which to operate.

    Returns
    -------
    reportnames : list of str
        A list with the names of all the reports in the Design
    """
    oModule = get_module(oDesign, "ReportSetup")
    report_list = list(oModule.GetAllReportNames())
    return map(str, report_list)

@conf.checkDefaultDesign
def add_traces(oDesign,
               report_name,
               setup_name,
               sweep_name,
               context_array,
               families_array,
               report_data_array):
    """
    Creates a new trace and adds it to the specified report.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS oDesign object upon which to operate.
    reportName : string
        The name of the report.
    SetupName : string
        Name of HFSS setup to use, for example "Setup1"
    SweepName : string
        Name of HFSS sweep to use, for example "LastAdaptive"
    ContextArray : list of strings
        Context for which the expression is being evaluated. This can be an empty string if there is no context.
        Must be passed in as a pair, for example
            ["Domain:=", "Sweep"]
            ["Domain:=", "Time"]
            ["Context:=", "Infinite Sphere"]
    FamiliesArray : list of strings
        Contains sweep definitions for the report.
        Must be passed in as a pair, for example
            ["VariableName:=", "Value"]
            ["Theta:=", "All")
    ReportDataArray : list of strings
        This array contains the report quantity and X, Y, and (Z) axis definitions.
        ["X Component:=", VariableName, "Y Component:=", VariableName]

    """
    check_setup(oDesign, setup_name)
    if sweep_name != "LastAdaptive":
        check_sweep(oDesign, setup_name, sweep_name)
    solution_name = setup_name + " : " + sweep_name

    oModule = get_module(oDesign, "ReportSetup")
    oModule.AddTraces(report_name, solution_name, context_array, families_array, report_data_array, [])

@conf.checkDefaultDesign
def rename_trace(oDesign, report_name, trace_name, new_name):
    """
    Rename a trace in a plot.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS oDesign object upon which to operate.
    reportName : string
        The name of the report.
    traceName : string
        The name of the trace to be renamed.
    newName : string
        The new name of the trace.
    Returns
    -------
    None
    """
    oModule = get_module(oDesign, "ReportSetup")
    oModule.RenameTrace(report_name, trace_name, new_name)

@conf.checkDefaultDesign
def change_report_properties(oDesign,
                            report_name,
                            tab_name,
                            props_path,
                            props_dict):
    """
    Change the properties of a report.

    NOTE: There are too many values for the different parameters depending on
    which properties are desired to be changed and the type of report, and the
    HFSS documentation is not very informative about them.
    The recomendation is to first record an script in HFSS changing the desired
    properties and then writing them here accordingly.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS oDesign object upon which to operate.
    reportName : string
        The name of the report.
    tab_name : string
        The name of the main tab where the properties are located.
        For example: 'General', 'Scaling' or 'Grid'.
    props_path : string
        The particular tab where the properties are located.
        For example: 'General', 'AxisX' or 'PolarGrid'
    props_dict : Dictionary
        Dictionary with the property names and their values.
        If the property is 'Color', a 3 RGB integer values list must be passed along.
        For example: {'Min Scale': "-20",
                      'Max Scale': "20",
                      'Spacing': "10",
                      'Auto Scale': False}
    Returns
    -------
    None
    """
    oModule = get_module(oDesign, "ReportSetup")

    changedProps_array = ["NAME:ChangedProps"]
    for key in props_dict:
        prop_list = ['NAME:'+str(key)]
        if str(key)=='Color':
            prop_list += ['R:=', props_dict[key][0],
                          'G:=', props_dict[key][1],
                          'B:=', props_dict[key][2]]
        else:
            prop_list += ['Value:=', str(props_dict[key])]
        changedProps_array.append(prop_list)

    properties_array = ["NAME:AllTabs",
                        ["NAME:"+tab_name,
                         ["NAME:PropServers", report_name+':'+props_path],
                         [changedProps_array]]]

    oModule.ChangeProperty(properties_array)

@conf.checkDefaultDesign
def check_setup(oDesign, setup_name):
    """
    Check that SetupName is in the current design. If not raise an exception.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS oDesign object upon which to operate.
    SetupName : str
        Name of HFSS setup to use, for example "Setup1"

    Returns
    -------
    None
    """
    # Get all of the setups in the design and test the name
    setups = get_setups(oDesign)
    if setup_name not in setups:
        raise Exception("SetupName not in design.")

@conf.checkDefaultDesign
def check_sweep(oDesign, setup_name, sweepname):
    """
    Check that SweepName is in the SetupName. If not raise an exception.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS oDesign object upon which to operate.
    SetupName : str
        Name of HFSS setup to use, for example "Setup1"
    SweepName : str
        Name of HFSS sweep to use, for example "LastAdaptive"

    Returns
    -------
    None
    """
    # Get all of the sweeps in the setup and test the name
    sweeps = get_sweeps(oDesign, setup_name)
    if sweepname not in sweeps:
        raise Exception("SweepName not in the Setup.")
