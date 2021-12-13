# -*- coding: utf-8 -*-
"""
Functions in this module correspond more or less to the functions described
in the HFSS Scripting Guide, Section "Boundary and Excitation Module Script
Commands".

At last count there were 4 functions implemented out of 20.
"""
from __future__ import division, print_function, unicode_literals, absolute_import
import re

from hycohanz.design import get_module
from hycohanz.expression import Expression as Ex

def assign_perfect_e(oDesign, boundaryname, facelist, InfGroundPlane=False):
    """
    Create a perfect E boundary.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    boundaryname : str
        The name to give this boundary in the Boundaries tree.
    facelist : list of ints
        The faces to assign to this boundary condition.

    Returns
    -------
    None
    """
    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    oBoundarySetupModule.AssignPerfectE(["Name:" + boundaryname, "Faces:=", facelist, "InfGroundPlane:=", InfGroundPlane])

def assign_radiation(oDesign,
                     faceidlist,
                     IsIncidentField=False,
                     IsEnforcedField=False,
                     IsFssReference=False,
                     IsForPML=False,
                     UseAdaptiveIE=False,
                     IncludeInPostproc=True,
                     Name='Rad1'):
    """
    Assign a radiation boundary on the given faces.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    faceidlist : list of ints
        The faces to assign to this boundary condition.
    IsIncidentField : bool
        If True, same as checking the "Incident Field" radio button in the
        Radiation Boundary setup dialog.  Mutually-exclusive with
        IsEnforcedField
    IsEnforcedField : bool
        If True, same as checking the "Enforced Field" radio button in the
        Radiation Boundary setup dialog.  Mutually-exclusive with
        IsIncidentField
    IsFssReference : bool
        If IsEnforcedField is False, is equivalent to checking the
        "Reference for FSS" check box i the Radiation Boundary setup dialog.
    IsForPML : bool
        Not explored at this time.  Likely use case is when defining a
        radiation boundary in conjuction with PMLs where the boundary lies on
        the surface between the PML and the PML base object.
    UseAdaptiveIE : bool
        Not explored at this time.  It is likely that setting this to True is
        equivalent to selecting the "Model exterior as HFSS-IE domain" check
        box in the Radiation Boundary setup dialog.
    IncludeInPostproc : bool
        Not explored at this time.  Likely use case is to remove certain
        boundaries from consideration during certain postprocessing
        operations, such as when computing the radiation pattern.
    Name : str
        The name to assign the boundary.
    Returns
    -------
    None
    """
    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    arg = ["NAME:{0}".format(Name),
           "Faces:=", faceidlist,
           "IsIncidentField:=", IsIncidentField,
           "IsEnforcedField:=", IsEnforcedField,
           "IsFssReference:=", IsFssReference,
           "IsForPML:=", IsForPML,
           "UseAdaptiveIE:=", UseAdaptiveIE,
           "IncludeInPostproc:=", IncludeInPostproc]

    oBoundarySetupModule.AssignRadiation(arg)

def assign_perfect_h(oDesign, boundaryname, facelist):
    """
    Create a perfect H boundary.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    boundaryname : str
        The name to give this boundary in the Boundaries tree.
    facelist : list of ints
        The faces to assign to this boundary condition.

    Returns
    -------
    None
    """
    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    oBoundarySetupModule.AssignPerfectH(["Name:" + boundaryname, "Faces:=", facelist])

def assign_waveport_multimode(oDesign,
                              faceidlist,
                              portname="WP1",
                              Nmodes=1,
                              RenormalizeAllTerminals=True,
                              UseLineAlignment=False,
                              DeembedDistance=0,
                              ShowReporterFilter=False,
                              ReporterFilter=[True],
                              UseAnalyticAlignment=False):
    """
    Assign a waveport excitation using multiple modes.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    portname : str
        Name of the port to create.
    faceidlist : list
        List of face id integers.
    Nmodes : int
        Number of modes with which to excite the port.

    Returns
    -------
    None
    """
    if DeembedDistance == 0:
        DoDeembed = False
    else:
        DoDeembed = True

    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    modesarray = ["NAME:Modes"]
    for n in range(0, Nmodes):
        modesarray.append(["NAME:Mode" + str(n + 1),
                           "ModeNum:=", n + 1,
                           "UseIntLine:=", False])

    waveportarray = ["NAME:" + portname,
                     "Faces:=", faceidlist,
                     "NumModes:=", Nmodes,
                     "RenormalizeAllTerminals:=", RenormalizeAllTerminals,
                     "UseLineModeAlignment:=", UseLineAlignment,
                     "DoDeembed:=", DoDeembed,
                     "DeembedDist:=", Ex(DeembedDistance).expr,
                     modesarray,
                     "ShowReporterFilter:=", ShowReporterFilter,
                     "ReporterFilter:=", ReporterFilter,
                     "UseAnalyticAlignment:=", UseAnalyticAlignment]

    oBoundarySetupModule.AssignWavePort(waveportarray)

def assign_waveport_intline(oDesign,
                              faceidlist,
                              startCoord,
                              stopCoord,
                              portname="WP1",
                              RenormalizeAllTerminals=True,
                              UseLineAlignment=False,
                              DeembedDistance=0,
                              AlignmentGroup=0,
                              CharImp="Zpi",
                              ShowReporterFilter=False,
                              ReporterFilter=[True],
                              UseAnalyticAlignment=False):
    """
    Assign a waveport excitation using one mode through an integration line.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    faceidlist : list
        List of face id integers.
    startCoord : list of floats or hycohanz Expression objects
        Integration line origin coordinates (x, y, z)
    stopCoord : list of floats or hycohanz Expression objects
        Integration line end coordinates (x, y, z)
    portname : str
        Name of the port to create.

    Returns
    -------
    None
    """
    if DeembedDistance == 0:
        DoDeembed = False
    else:
        DoDeembed = True

    # For some reason, HFSS does not accept integration line coordinates
    # in the common HFSS expression form nor without units, so the expression
    # needs to be evaluated before feeding it to the program
    for i in range(3):
        units = re.search(r'[a-zA-Z]+', Ex(startCoord[i]).expr)
        startCoord[i] = str(eval(re.sub("[a-zA-Z]+", "", Ex(startCoord[i]).expr)))
        if units != None:
            startCoord[i] = startCoord[i] + units[0]
        else:
            startCoord[i] = startCoord[i] + "meter"
        units = re.search(r'[a-zA-Z]+', Ex(stopCoord[i]).expr)
        stopCoord[i] = str(eval(re.sub("[a-zA-Z]+", "", Ex(stopCoord[i]).expr)))
        if units != None:
            stopCoord[i] = stopCoord[i] + units[0]
        else:
            stopCoord[i] = stopCoord[i] + "meter"

    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    modesarray = ["NAME:Modes"]
    modesarray.append(["NAME:Mode1",
                        "ModeNum:=", 1,
                        "UseIntLine:=", True,
                        [
                            "NAME:IntLine",
                            "Start:="       , startCoord,
                            "End:="         , stopCoord
                        ],
                        "AlignmentGroup:=", AlignmentGroup,
					    "CharImp:=", CharImp])

    waveportarray = ["NAME:" + portname,
                     "Faces:=", faceidlist,
                     "NumModes:=", 1,
                     "RenormalizeAllTerminals:=", RenormalizeAllTerminals,
                     "UseLineModeAlignment:=", UseLineAlignment,
                     "DoDeembed:=", DoDeembed,
                     "DeembedDist:=", Ex(DeembedDistance).expr,
                     modesarray,
                     "ShowReporterFilter:=", ShowReporterFilter,
                     "ReporterFilter:=", ReporterFilter,
                     "UseAnalyticAlignment:=", UseAnalyticAlignment]

    oBoundarySetupModule.AssignWavePort(waveportarray)

def assign_lumpedport_intline(oDesign,
                              faceidlist,
                              startCoord,
                              stopCoord,
                              portname="LP1",
                              RenormalizeAllTerminals=True,
                              DeembedDistance=0,
                              AlignmentGroup=0,
                              CharImp="Zpi",
                              Impedance="50ohm",
                              ShowReporterFilter=False,
                              ReporterFilter=[True]):
    """
    Assign a lumped port excitation using one mode through an integration line.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    faceidlist : list
        List of face id integers.
    startCoord : list of floats or hycohanz Expression objects
        Integration line origin coordinates (x, y, z)
    stopCoord : list of floats or hycohanz Expression objects
        Integration line end coordinates (x, y, z)
    portname : str
        Name of the port to create.

    Returns
    -------
    None
    """
    if DeembedDistance == 0:
        DoDeembed = False
    else:
        DoDeembed = True

    # Debuaging
    # print([Ex(startCoord[0]).expr, Ex(startCoord[1]).expr, Ex(startCoord[2]).expr])
    # print([Ex(stopCoord[0]).expr, Ex(stopCoord[1]).expr, Ex(stopCoord[2]).expr])

    # For some reason, HFSS does not accept integration line coordinates
    # in the common HFSS expression form nor without units, so the expression
    # needs to be evaluated before feeding it to the program
    for i in range(3):
        units = re.search(r'[a-zA-Z]+', Ex(startCoord[i]).expr)
        startCoord[i] = str(eval(re.sub("[a-zA-Z]+", "", Ex(startCoord[i]).expr)))
        if units != None:
            startCoord[i] = startCoord[i] + units[0]
        else:
            startCoord[i] = startCoord[i] + "meter"
        units = re.search(r'[a-zA-Z]+', Ex(stopCoord[i]).expr)
        stopCoord[i] = str(eval(re.sub("[a-zA-Z]+", "", Ex(stopCoord[i]).expr)))
        if units != None:
            stopCoord[i] = stopCoord[i] + units[0]
        else:
            stopCoord[i] = stopCoord[i] + "meter"

    # Debuaging
    # print([Ex(startCoord[0]).expr, Ex(startCoord[1]).expr, Ex(startCoord[2]).expr])
    # print([Ex(stopCoord[0]).expr, Ex(stopCoord[1]).expr, Ex(stopCoord[2]).expr])

    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    modesarray = ["NAME:Modes"]
    modesarray.append(["NAME:Mode1",
                        "ModeNum:=", 1,
                        "UseIntLine:=", True,
                        [
                            "NAME:IntLine",
                            "Start:=", startCoord,
                            "End:=", stopCoord
                        ],
                        "AlignmentGroup:=", AlignmentGroup,
					    "CharImp:=", CharImp,
                        "RenormImp:=", Impedance])

    lumpedportarray = ["NAME:" + portname,
                     "Faces:=", faceidlist,
                     "DoDeembed:=", DoDeembed,
                     "DeembedDist:=", Ex(DeembedDistance).expr,
                     "RenormalizeAllTerminals:=", RenormalizeAllTerminals,
                     modesarray,
                     "ShowReporterFilter:=", ShowReporterFilter,
                     "ReporterFilter:=", ReporterFilter,
                     "Impedance:=", Impedance]

    oBoundarySetupModule.AssignLumpedPort(lumpedportarray)
