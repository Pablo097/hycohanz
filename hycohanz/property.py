# -*- coding: utf-8 -*-
"""
HFSS functions that use the Property module. Functions in this module correspond
more or less to the functions described in the HFSS Scripting Guide,
Section "Property Script Commands".

At last count there were 1 functions implemented out of 7.
"""
from __future__ import division, print_function, unicode_literals, absolute_import

import hycohanz.conf as conf
from hycohanz.expression import Expression
import re
import math
from quantiphy import Quantity

@conf.checkDefaultDesign
def add_properties(oDesign, name, value):
    """
    Add design properties.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design from which to retrieve the module.
    name : list of str
        The name of the properties to add.
    value : list of Hyphasis Expression object
        The values of the properties.

    Returns
    -------
    None

    """
    propserversarray = ["NAME:PropServers", "LocalVariables"]

    newpropsarray = ["NAME:NewProps"]

    for n in range(len(name)):
        if isinstance(value[n], list):
            valueAux = '['+','.join([Expression(i).expr for i in value[n]])+']'
        else:
            valueAux = Expression(value[n]).expr

        newpropsarray.append(["NAME:" + name[n],
                           "PropType:=", "VariableProp",
                           "UserDef:=", True,
                           "Value:=", valueAux])

    proptabarray = ["NAME:LocalVariableTab", propserversarray, newpropsarray]

    oDesign.ChangeProperty(["NAME:AllTabs", proptabarray])

@conf.checkDefaultDesign
def add_property(oDesign, name, value):
    """
    Add a design property.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design from which to retrieve the module.
    name : str
        The name of the property to add.
    value : Hyphasis Expression object
        The value of the property.

    Returns
    -------
    None

    """
    add_properties(oDesign, [name], [value])

@conf.checkDefaultProject
def set_variable(oProject, name, value):
    """
    Change a design property.  This function differs significantly from
    SetVariableValue() in that it makes the reasonable assumption that
    if the variable contains '$', then the variable is global; otherwise,
    it is assumed to be a local variable.

    Parameters
    ----------
    oProject : pywin32 COMObject
        The HFSS design from which to retrieve the module.
    name : str
        The name of the property/variable to edit.
    value : Hyphasis Expression object
        The new value of the property.

    Returns
    -------
    None

    """
    if '$' in name:
        oProject.SetVariableValue(name,Expression(value).expr)
    else:
        oDesign = oProject.GetActiveDesign()
        oDesign.SetVariableValue(name,Expression(value).expr)

@conf.checkDefaultProject
def get_variables(oProject,oDesign=''):
    """
    get list of non-indexed variables.

    Parameters
    ----------
    oProject : pywin32 COMObject
        The HFSS design from which to retrieve the variables.
    oDesign : pywin32 COMObject
        Optional, if specified function returns variable list of oDesign.


    Returns
    -------
    variable_list: list of str
    list of non-indexed project/design variables

    """
    if oDesign=='':
        variable_list = list(oProject.GetVariables())
    else:
        variable_list = list(oDesign.GetVariables())
    return map(str,variable_list)

@conf.checkDefaultDesign
def get_variable_value(oDesign,varName):
    """
    get the value of a variable.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design from which to retrieve the variable.
    varName : str
        string with the name of the variable

    Returns
    -------
    variable_value: str
    string representing the value of the variable

    """
    return oDesign.GetVariableValue(Expression(varName).expr)

@conf.checkDefaultDesign
def expand_expression(oDesign, exprValue):
    """
    Expands an expression with all the numeric values and
    their units from HFSS variables

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design from which to retrieve the variable values.
    exprValue : int, float, str or hycohanz Expression object
        expression to expand

    Returns
    -------
    str1: str
        string representing the expression expanded as numbers and units only

    Example Usage
    -------------
    # Assume the HFSS oDesign has the following variables declared in the project
    # > varA = 18um
    # > varB = 4GHz
    # > varC = c0/varB + varA
    >>> expand_expression(oDesign, 'varC')
    '299792458.0/4GHz + 18um'
    """
    str1 = Expression(exprValue).expr
    # print('Expresion a sustituir: '+str1)
    variablelist = re.findall(r'\b[a-zA-Z]\w*', str1)
    # print(variablelist)
    if len(variablelist) > 0:
    	for variable in variablelist:
    		# print('Variable a sustituir: '+variable)
            if variable=="c0":
                str1 = str1.replace(variable, str(float(Quantity('c'))))
            elif variable=="e0":
                str1 = str1.replace(variable, str(float(Quantity('eps0'))))
            elif variable=="u0":
                str1 = str1.replace(variable, str(float(Quantity('mu0'))))
            elif variable=="pi":
                str1 = str1.replace(variable, str(math.pi))
            else:
                str2 = get_variable_value(oDesign, variable)
                # print('De HFSS: '+variable+' = '+str2)
                str1 = str1.replace(variable, expand_expression(oDesign, str2))
    # print('Expresion expandida: '+str1)
    return str1

@conf.checkDefaultDesign
def eval_expression(oDesign, exprValue):
    """
    Evaluates an expression taking the necessary HFSS design variables

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design from which to retrieve the possible variable values.
    exprValue : int, float, str or hycohanz Expression object
        expression to evaluate

    Returns
    -------
    value: float
        evaluated value in SI units

    Example Usage
    -------------
    # Assume the HFSS oDesign has the following variables declared in the project
    # > varA = 18um
    # > varB = 4GHz
    # > varC = c0/varB + varA
    >>> expand_expression(oDesign, 'varC')
    0.0749661145
    """
    str1 = expand_expression(oDesign, exprValue)
    # print(str1)
    for variableWithUnits in list(set(re.findall(r'\b[\d]+\.?[\d]*[A-Za-z]+', str1))):
        # print(variableWithUnits)
        if 'mil' in variableWithUnits:	# Catch non-standard units:
            aux = variableWithUnits.replace('mil','')
            str1 = str1.replace(variableWithUnits, str(float(aux)*2.54e-5))
        else:
            str1 = str1.replace(variableWithUnits, str(float(Quantity(variableWithUnits))))
        # print(str1)
    value = eval(str1)
    # print('Expresion con unidades en SI: '+str1+' = '+str(value))
    return value
