# -*- coding: utf-8 -*-
"""
HFSS Fields Calculator functions. Functions in this module correspond more or
less to the functions described in the HFSS Scripting Guide, Section "Fields
Calculator Script Commands".

At last count there were 7 functions implemented out of 28.
"""

from __future__ import division, print_function, unicode_literals, absolute_import

import hycohanz.conf as conf
from hycohanz.design import get_module
from hycohanz.expression import Expression as Ex

@conf.checkDefaultDesign
def enter_vol(oDesign, VolumeName):
    """
    Enters a volume defined in the 3D Modeler editor into the Fields Calculator.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS oDesign object upon which to operate.
    VolumeName : str
        Name of a volume defined in the 3D Modeler editor.

    Returns
    -------
    None
    """
    oFieldsReporter = get_module(oDesign, 'FieldsReporter')
    oFieldsReporter.EnterVol(VolumeName)

@conf.checkDefaultDesign
def calc_op(oDesign, OperationString):
    """
    Performs a calculator operation in the Fields Calculator.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS oDesign object upon which to operate.
    OperationString : str
        The text on the corresponding calculator button.

    Returns
    -------
    None
    """
    oFieldsReporter = get_module(oDesign, 'FieldsReporter')
    oFieldsReporter.CalcOp(OperationString)

@conf.checkDefaultDesign
def clc_eval(oDesign, setupname, sweepname, freq, phase, variablesdict):
    """
    Evaluates the expression at the top of the Fields Calculator stack using
    the provided solution name and variable values.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS oDesign object upon which to operate.
    setupname : str
        Name of HFSS setup to use, for example "Setup1"
    sweepname : str
        Name of HFSS sweep to use, for example "LastAdaptive"
    freq : float
        Frequency at which to evaluate the calculator expression.
    phase : float
        Phase in degrees at which to evaluate the calculator expression.
    variablesdict : dict
        Dictionary listing the variables and their values that define the
        design variation, except for 'Freq' and 'Phase'.  Variable names are the keys (strings), and the values are
        the values of the variable (floats). For example: {'radius': 0.5, 'height': '2.0'}

    Returns
    -------
    None
    """
    solutionname = setupname + " : " + sweepname

    variablesarray = ["Freq:=", str(freq) + 'Hz', "Phase:=", str(phase) + 'deg']

    for key in variablesdict:
        variablesarray += [str(key) + ':=', str(variablesdict[key])]

    print('solutionname: ' + str(solutionname))
    print('variablesarray: ' + str(variablesarray))

    oFieldsReporter = get_module(oDesign, 'FieldsReporter')
    oFieldsReporter.ClcEval(solutionname, variablesarray)

@conf.checkDefaultDesign
def enter_qty(oDesign, FieldQuantityString):
    """
    Enters a field quantity into the Fields Calculator.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS oDesign object upon which to operate.
    FieldQuantityString : str
        The field quantity to be entered onto the stack.

    Returns
    -------
    None
    """
    oFieldsReporter = get_module(oDesign, 'FieldsReporter')
    return oFieldsReporter.EnterQty(FieldQuantityString)

@conf.checkDefaultDesign
def copy_named_expr_to_stack(oDesign, ExpressionString):
    """
    Enters an expression existing in the Fields Calculator into the stack.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS oDesign object upon which to operate.
    ExpressionString : str
        The field calculator expression to be entered onto the stack.
        For example, "ComplexMag_E" or "Vector_RealPoynting"

    Returns
    -------
    None
    """
    oFieldsReporter = get_module(oDesign, 'FieldsReporter')
    return oFieldsReporter.CopyNamedExprToStack(ExpressionString)

@conf.checkDefaultDesign
def add_named_expr(oDesign, ExpressionString):
    """
    Creates a named expression using the expression at the top of the stack.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS oDesign object upon which to operate.
    ExpressionString : str
        The name of the field calculator expression to be created.
        For example, "My_new_E"

    Returns
    -------
    None
    """
    oFieldsReporter = get_module(oDesign, 'FieldsReporter')
    return oFieldsReporter.AddNamedExpr(ExpressionString)

@conf.checkDefaultDesign
def get_top_entry_value(oDesign, setupname, sweepname, freq, phase, variablesdict):
    """
    Evaluates the expression at the top of the stack using the provided
    solution name and variable values.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS oDesign object upon which to operate.
    SolutionName : str
        Name of solution to use, for example "Setup1 : LastAdaptive"
    Freq : float
        Frequency at which to evaluate the calculator expression.
    Phase : float
        Phase in degrees at which to evaluate the calculator expression.
    variablesdict : dict
        Dictionary listing the variables and their values that define the
        design variation, except for 'Freq' and 'Phase'.  Variable names are the keys (strings), and the values are
        the values of the variable (floats). For example: {'radius': 0.5, 'height': '2.0'}

    Returns
    -------
    result : str
        A string representation of the value at the top of the Calculator stack.

    """
    solutionname = setupname + " : " + sweepname

    variablesarray = ["Freq:=", str(freq) + 'Hz', "Phase:=", str(phase) + 'deg']

    for key in variablesdict:
        variablesarray += [str(key) + ':=', str(variablesdict[key])]

    print('solutionname: ' + str(solutionname))
    print('variablesarray: ' + str(variablesarray))

    oFieldsReporter = get_module(oDesign, 'FieldsReporter')
    return oFieldsReporter.GetTopEntryValue(solutionname, variablesarray)

@conf.checkDefaultDesign
def export_on_grid(oDesign,
                    filename,
                    startCoord,
                    stopCoord,
                    step,
                    setupName,
                    freq,
                    phase = 0,
                    variablesDict = {},
                    includePoints = True,
                    coordType = 'Cartesian',
                    offsetCoord = [0, 0, 0]
                    ):
    """
    Exports the expression at the top of the Fields Calculator stack using
    the provided solution name and variable values.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS oDesign object upon which to operate.
    filename : str
        Name of the output file. Must include full path and end with '.fld'
    startCoord : list of floats or hycohanz Expressions
        Minimum values for the coordinate components of the grid system.
    stopCoord : list of floats or hycohanz Expressions
        Maximum values for the coordinate components of the grid system.
    step : list of floats or hycohanz Expressions
        Spacing values for the coordinate components of the grid system.
    setupname : str
        Name of HFSS setup to use, for example "Setup1"
    freq : float
        Frequency at which to evaluate the calculator expression.
    phase : float
        Phase in degrees at which to evaluate the calculator expression.
    variablesdict : dict
        Dictionary listing the variables and their values that define the
        design variation, except for 'Freq' and 'Phase'.  Variable names are the keys (strings), and the values are
        the values of the variable (floats). For example: {'radius': 0.5, 'height': '2.0'}
    includePoints : boolean
        Specifies whether include points in the output file.
    coordType : string
        Type of coordinate system. "Cartesian" (default) | "Cylindrical" | "Spherical".
    offsetCoord : list of floats or hycohanz Expressions
        Origin for the offset coordinate system.

    Returns
    -------
    None
    """
    if filename[-4:] != '.fld':
        filename = filename + '.fld'

    startCoord = [Ex(coord).expr for coord in startCoord]
    stopCoord = [Ex(coord).expr for coord in stopCoord]
    step = [Ex(coord).expr for coord in step]
    offsetCoord = [Ex(coord).expr for coord in offsetCoord]

    variablesarray = ["Freq:=", str(freq) + 'Hz', "Phase:=", str(phase) + 'deg']
    for key in variablesdict:
        variablesarray += [str(key) + ':=', str(variablesdict[key])]

    oFieldsReporter = get_module(oDesign, 'FieldsReporter')
    oFieldsReporter.ExportOnGrid(filename, startCoord, stopCoord, step,
                            setupName, variablesarray,
                            includePoints, coordType, offsetCoord)



if __name__ == "__main__":
    import doctest
    doctest.testmod()
