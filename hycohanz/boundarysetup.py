# -*- coding: utf-8 -*-
"""
Functions in this module correspond more or less to the functions described
in the HFSS Scripting Guide, Section "Boundary and Excitation Module Script
Commands".

At last count there were 6 functions implemented out of 20.
"""
from __future__ import division, print_function, unicode_literals, absolute_import
import re

import hycohanz.conf as conf
from hycohanz.design import get_module
from hycohanz.property import eval_expression
from hycohanz.expression import Expression as Ex

@conf.checkDefaultDesign
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

@conf.checkDefaultDesign
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

@conf.checkDefaultDesign
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

@conf.checkDefaultDesign
def assign_anisotropic_impedance(oDesign,
                                faceList,
                                resistanceArray,
                                reactanceArray,
                                boundaryName = "AnisoImp1",
                                UseInfiniteGP = False):
    """
    Create an anisotropic impedance boundary.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    faceList : list of ints
        The faces to assign this boundary condition.
    resistanceArray : list of floats and/or hycohanz Expression objects
    reactanceArray : list of floats and/or hycohanz Expression objects
        Lists with the anisotropic resistance and reactance values
        ordered as [Zxx, Zxy, Zyx, Zyy]

    Returns
    -------
    None
    """
    oModule = get_module(oDesign, "BoundarySetup")

    anisImpArray = ["NAME:"+boundaryName,
                    "Faces:="		, faceList,
                    "UseInfiniteGroundPlane:=", False,
                    "CoordSystem:="		, "Global",
                    "HasExternalLink:="	, False,
                    "ZxxResistance:="	, Ex(resistanceArray[0]).expr,
                    "ZxxReactance:="	, Ex(reactanceArray[0]).expr,
                    "ZxyResistance:="	, Ex(resistanceArray[1]).expr,
                    "ZxyReactance:="	, Ex(reactanceArray[1]).expr,
                    "ZyxResistance:="	, Ex(resistanceArray[2]).expr,
                    "ZyxReactance:="	, Ex(reactanceArray[2]).expr,
                    "ZyyResistance:="	, Ex(resistanceArray[3]).expr,
                    "ZyyReactance:="	, Ex(reactanceArray[3]).expr]
    oModule.AssignAnisotropicImpedance(anisImpArray)

@conf.checkDefaultDesign
def assign_waveport(oDesign,
                      faceidlist,
                      portname="WP1",
                      Nmodes=1,
                      startCoordsList=[],
                      stopCoordsList=[],
                      RenormalizeAllTerminals=False,
                      Impedance=50,
                      UseLineAlignment=False,
                      DeembedDistance=0,
                      AlignmentGroup=None,
                      CharImp=None,
                      ShowReporterFilter=False,
                      ReporterFilter=None,
                      UseAnalyticAlignment=False):
    """
    Assign a waveport excitation. Several modes can be defined, and integration
    lines for all or the first n modes can be specified.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    faceidlist : list
        List of face id integers.
    portname : str
        Name of the port to create.
    Nmodes : int
        Number of modes with which to excite the port.
    startCoordsList : list of list of floats or hycohanz Expression objects
        Integration line origin coordinates (x, y, z) for each mode
        which needs them. If the number of 3-coordinates lists are less than
        Nmodes, the rest of the modes will be excited without integration lines
    stopCoordsList : list of list of floats or hycohanz Expression objects
        Integration line end coordinates (x, y, z) for each mode
        which needs them, as with startCoordList
    Impedance : int
        Renormalization impedance value in Ohms.
    CharImp : list of str
        The type of characteristic impedance (Z0) to use for each mode.
        Can be 'Zpi', 'Zpv' or 'Zvi'. If the length of this list is smaller
        than Nmodes, 'Zpi' is used for the rest of the modes.

    Returns
    -------
    None
    """
    if DeembedDistance == 0:
        DoDeembed = False
    else:
        DoDeembed = True

    # Obtain coordinates and their values
    if startCoordsList:
        numIntLines = len(startCoordsList)
        if numIntLines != len(stopCoordsList):
            print("Error in assign_waveport: The lists of start and stop coordinates must have the same length.\n")
        startCoords = []
        stopCoords = []
        if type(startCoordsList[0]) != list:
            if numIntLines == 3:
                # Assume that coordinates for only one mode have been given
                startCoords.append([str(eval_expression(oDesign, aux)*1e3)+'mm' for aux in startCoordsList])
                stopCoords.append([str(eval_expression(oDesign, aux)*1e3)+'mm' for aux in stopCoordsList])
                numIntLines = 1
            else:
                print("Error in assign_waveport: Wrong coordinates format (need list of 3-coordinates lists).\n")
        else:
            for index in range(numIntLines):
                # For some reason, HFSS does not accept integration line coordinates
                # in the common HFSS expression form nor without units, so the expression
                # needs to be evaluated before feeding it to the program
                startCoords.append([str(eval_expression(oDesign, aux)*1e3)+'mm' for aux in startCoordsList[index]])
                stopCoords.append([str(eval_expression(oDesign, aux)*1e3)+'mm' for aux in stopCoordsList[index]])
    else:
        numIntLines = 0

    # Make sure CharImp has length equal to the number of modes
    if CharImp:
        if len(CharImp)<Nmodes:
            CharImp += ["Zpi"]*(Nmodes-len(CharImp))
    else:
        CharImp = ["Zpi"]*Nmodes

    # Fill incomplete required fields
    if not ReporterFilter:
        ReporterFilter = [True]*Nmodes

    # Correctly format the impedance if needed
    if RenormalizeAllTerminals:
        if Ex(Impedance).expr[-3:]!='ohm':
            Impedance = f"{Impedance}ohm"

    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")

    modesarray = ["NAME:Modes"]
    for n in range(0, Nmodes):
        oneModeArray = ["NAME:Mode" + str(n + 1),
                        "ModeNum:=", n + 1]
        if n<numIntLines:
            intLineArray = ["NAME:IntLine",
                            "Start:=", startCoords[n],
                            "End:=", stopCoords[n]]
            oneModeArray += ["UseIntLine:=", True, intLineArray]
        else:
            oneModeArray += ["UseIntLine:=", False]
        if UseLineAlignment and AlignmentGroup and n<len(AlignmentGroup):
            oneModeArray += ["AlignmentGroup:=", AlignmentGroup[n]]
        oneModeArray += ["CharImp:=", CharImp[n]]
        if RenormalizeAllTerminals:
            oneModeArray += ["RenormImp:=", Impedance]
        modesarray.append(oneModeArray)

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

@conf.checkDefaultDesign
def assign_lumpedport(oDesign,
                      faceidlist,
                      startCoord,
                      stopCoord,
                      portname="LP1",
                      RenormalizeAllTerminals=True,
                      PortImpedance=50,
                      DeembedDistance=0,
                      AlignmentGroup=0,
                      CharImp="Zpi",
                      RenormImpedance=50,
                      ShowReporterFilter=False,
                      ReporterFilter=[True]):
    """
    Assign a lumped port excitation using one mode through an integration line.
    (This is the only allowed option for lumped ports in HFSS)

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

    # Correctly format the impedance if needed
    if Ex(PortImpedance).expr[-3:]!='ohm':
        PortImpedance = f"{PortImpedance}ohm"
    if Ex(RenormImpedance).expr[-3:]!='ohm':
        RenormImpedance = f"{RenormImpedance}ohm"

    # Debuaging
    # print([Ex(startCoord[0]).expr, Ex(startCoord[1]).expr, Ex(startCoord[2]).expr])
    # print([Ex(stopCoord[0]).expr, Ex(stopCoord[1]).expr, Ex(stopCoord[2]).expr])

    # For some reason, HFSS does not accept integration line coordinates
    # in the common HFSS expression form nor without units, so the expression
    # needs to be evaluated before feeding it to the program
    startCoord = [str(eval_expression(oDesign, aux)*1e3)+'mm' for aux in startCoord]
    stopCoord = [str(eval_expression(oDesign, aux)*1e3)+'mm' for aux in stopCoord]

    ## Old method
    # for i in range(3):
    #     units = re.search(r'[a-zA-Z]+', Ex(startCoord[i]).expr)
    #     startCoord[i] = str(eval(re.sub("[a-zA-Z]+", "", Ex(startCoord[i]).expr)))
    #     if units != None:
    #         startCoord[i] = startCoord[i] + units[0]
    #     else:
    #         startCoord[i] = startCoord[i] + "meter"
    #     units = re.search(r'[a-zA-Z]+', Ex(stopCoord[i]).expr)
    #     stopCoord[i] = str(eval(re.sub("[a-zA-Z]+", "", Ex(stopCoord[i]).expr)))
    #     if units != None:
    #         stopCoord[i] = stopCoord[i] + units[0]
    #     else:
    #         stopCoord[i] = stopCoord[i] + "meter"

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
                        "RenormImp:=", RenormImpedance])

    lumpedportarray = ["NAME:" + portname,
                     "Faces:=", faceidlist,
                     "DoDeembed:=", DoDeembed,
                     "DeembedDist:=", Ex(DeembedDistance).expr,
                     "RenormalizeAllTerminals:=", RenormalizeAllTerminals,
                     modesarray,
                     "ShowReporterFilter:=", ShowReporterFilter,
                     "ReporterFilter:=", ReporterFilter,
                     "Impedance:=", PortImpedance]

    oBoundarySetupModule.AssignLumpedPort(lumpedportarray)

@conf.checkDefaultDesign
def assign_current(oDesign,
                   objectsList,
                   startCoordsList,
                   endCoordsList,
                   name="Current1"):
    """
    Assign a current excitation.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    objectsList : list of strings
        List of the object names to which to apply the current excitation.
    startCoordsList : list of floats or hycohanz Expression objects
        Current line origin coordinates (x, y, z).
    endCoordsList : list of floats or hycohanz Expression objects
        Current line end coordinates (x, y, z).
    name : string
        Name of the current excitation to assign.

    Returns
    -------
    None
    """
    startCoords = [str(eval_expression(oDesign, aux)*1e3)+'mm' for aux in startCoordsList]
    endCoords = [str(eval_expression(oDesign, aux)*1e3)+'mm' for aux in endCoordsList]

    directionArray = ["NAME:Direction",
                      "Start:=", startCoords,
                      "End:=", endCoords]
    paramArray = ["NAME:"+name,
		          "Objects:=", objectsList,
		          directionArray]

    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    oBoundarySetupModule.AssignCurrent(paramArray)

@conf.checkDefaultDesign
def create_PML(oDesign,
                PMLObjectId,
                BaseObjectId,
                Thickness = "1mm",
                CreateJoiningObjs = False,
                Orientation="XAxis",
                UseFreq = True,
                MinFreq = "1GHz",
                MinBeta = 20,
                RadDist = "55.5865182541667mm"):
    """
    Create a PML (Perfectly Matched Load) boundary on the given object.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    PMLObjectId : int
        The ID of the object in which to create the PML. The object name must
        start with "PML", otherwise, its name will be automatically changed to
        include that prefix.
    BaseObjectId : int
        The ID of the object touching the PML object.
    Thickness : string or hycohanz Expression
        Thickness of the PML.
    Orientation : string
        Orientation of the PML. Can be "XAxis", "YAxis" or "ZAxis".
    Returns
    -------
    None
    """
    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    arg = ["NAME:PMLCreationSettings",
    		"UserDrawnGroup:=", True,
    		"PMLFaces:=", [],
    		"Thickness:=", Ex(Thickness).expr,
    		"CreateJoiningObjs:=", CreateJoiningObjs,
    		"PMLObj:=", PMLObjectId,
    		"BaseObj:=", BaseObjectId,
    		"Orientation:=", Orientation,
    		"UseFreq:=", UseFreq,
    		"MinFreq:=", Ex(MinFreq).expr,
    		"MinBeta:=", MinBeta,
    		"RadDist:=", Ex(RadDist).expr]

    oBoundarySetupModule.CreatePML(arg)

@conf.checkDefaultDesign
def assign_primary(oDesign,
                    faceidlist,
                    startCoord,
                    stopCoord,
                    boundaryName="Primary1",
                    reverseV=False):
    """
    Assign a primary periodic condition.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    faceidlist : list
        List of face id integers.
    startCoord : list of floats or hycohanz Expression objects
        Origin coordinates (x, y, z) of the periodicity vector U
    stopCoord : list of floats or hycohanz Expression objects
        End coordinates (x, y, z) of the periodicity vector U
    boundaryName : str
        Name of the bondary to create.
    reverseV: boolean
        Whether to reverse the direcion of the V vector perpendicular to U

    Returns
    -------
    None
    """
    startCoord = [str(eval_expression(oDesign, aux)*1e3)+'mm' for aux in startCoord]
    stopCoord = [str(eval_expression(oDesign, aux)*1e3)+'mm' for aux in stopCoord]

    oModule = get_module(oDesign, "BoundarySetup")
    coordSysArray = ["NAME:CoordSysVector",
                     "Origin:=", startCoord,
                     "UPos:=", stopCoord]
    primaryBondArray = ["NAME:" + boundaryName,
		                "Faces:=", faceidlist,
                        coordSysArray,
                        "ReverseV:=", reverseV]

    oModule.AssignPrimary(primaryBondArray)

@conf.checkDefaultDesign
def assign_secondary(oDesign,
                    faceidlist,
                    startCoord,
                    stopCoord,
                    primaryName,
                    boundaryName="Secondary1",
                    reverseV=False,
                    phi=None,
                    theta=None,
                    phase=None):
    """
    Assign a secondary periodic condition linked to a existing primary one.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    faceidlist : list
        List of face id integers.
    startCoord : list of floats or hycohanz Expression objects
        Origin coordinates (x, y, z) of the periodicity vector U
    stopCoord : list of floats or hycohanz Expression objects
        End coordinates (x, y, z) of the periodicity vector U
    primaryName : str
        Name of the primary periodic condition to link this secondary one
    boundaryName : str
        Name of the bondary to create.
    phi, theta, phase : float (in deg) or hycohanz Expression
        Scan angles (phi, theta) or Input Phase Delay (phase). Only one of
        the two can be used.
    reverseV: boolean
        Whether to reverse the direcion of the V vector perpendicular to U

    Returns
    -------
    None
    """
    startCoord = [str(eval_expression(oDesign, aux)*1e3)+'mm' for aux in startCoord]
    stopCoord = [str(eval_expression(oDesign, aux)*1e3)+'mm' for aux in stopCoord]

    if phi is not None:
        if type(phi)==int or type(phi)==float:
            phi = f"{phi}deg"
        if type(theta)==int or type(theta)==float:
            theta = f"{theta}deg"
        anglesArray = ["UseScanAngles:=", True,
                       "Phi:=", Ex(phi).expr,
                       "Theta:=", Ex(theta).expr]
    elif phase is not None:
        if type(phase)==int or type(phase)==float:
            phase = f"{phase}deg"
        anglesArray = ["UseScanAngles:=", False,
                       "Phase:=", Ex(phase).expr]
    else:
        anglesArray = ["UseScanAngles:=", True,
                       "Phi:=", "0deg",
                       "Theta:=", "0deg"]

    oModule = get_module(oDesign, "BoundarySetup")
    coordSysArray = ["NAME:CoordSysVector",
                     "Origin:=", startCoord,
                     "UPos:=", stopCoord]
    secondaryBondArray = ["NAME:" + boundaryName,
		                "Faces:=", faceidlist,
                        coordSysArray,
                        "ReverseV:=", reverseV,
                        "Primary:=", primaryName]
    secondaryBondArray += anglesArray

    oModule.AssignSecondary(secondaryBondArray)

@conf.checkDefaultDesign
def assign_floquetport(oDesign,
                      faceidlist,
                      portname="FP1",
                      Nmodes=2,
                      startCoordsA=[],
                      endCoordsA=[],
                      startCoordsB=[],
                      endCoordsB=[],
                      RenormalizeAllTerminals=True,
                      DeembedDistance=0,
                      ModesList = [{'IndexM':0, 'IndexN':0,
                                    'Polarization':'TE',
                                    'Attenuation':0,
                                    'AffectsRefinement':True},
                                   {'IndexM':0, 'IndexN':0,
                                    'Polarization':'TM',
                                    'Attenuation':0,
                                    'AffectsRefinement':True}],
                      ShowReporterFilter=False,
                      UseIntLine=False,
                      CharImp="Zpi",
                      UseScanAngles=True,
                      Phi=0,
                      Theta=0):
    """
    Assign a Floquet Port excitation. Several modes can be defined, each with
    its corresponding configuration dictionary.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design to which this function is applied.
    faceidlist : list
        List of face id integers.
    portname : str
        Name of the port to create.
    Nmodes : int
        Number of modes with which to excite the port.
    startCoordsA : list of floats or hycohanz Expression objects
    endCoordsA : list of floats or hycohanz Expression objects
    startCoordsB : list of floats or hycohanz Expression objects
    endCoordsB : list of floats or hycohanz Expression objects
        Integration line origin and end coordinates (x, y, z) for lattice
        vector A and B, which define the periodic lattice geometry.
    DeembedDistance : floats or hycohanz Expression
        Distance for deembeding the port.
    ModesList : list of dict
        List with dictionaries of the properties of each mode

    Returns
    -------
    None
    """
    if DeembedDistance == 0:
        DoDeembed = False
    else:
        DoDeembed = True

    # Obtain coordinates and their values
    startCoordsA = [str(eval_expression(oDesign, aux)*1e3)+'mm' for aux in startCoordsA]
    endCoordsA = [str(eval_expression(oDesign, aux)*1e3)+'mm' for aux in endCoordsA]
    startCoordsB = [str(eval_expression(oDesign, aux)*1e3)+'mm' for aux in startCoordsB]
    endCoordsB = [str(eval_expression(oDesign, aux)*1e3)+'mm' for aux in endCoordsB]

    if ((type(ModesList) == list and Nmodes != len(ModesList)) or
        (type(ModesList) == dict and Nmodes != 1)):
        print("Error in assign_floquetport: 'ModesList' must have Nmodes dictionaries.\n")
        return

    modesListArray = ["NAME:ModesList"]
    modesarray = ["NAME:Modes"]
    for n in range(0, Nmodes):
        modesarray.append(["NAME:Mode" + str(n + 1),
                            "ModeNum:=", n + 1,
                            "UseIntLine:=", UseIntLine,
                            "CharImp:=", CharImp])
        modesListArray.append(["NAME:Mode",
                               "ModeNumber:=", n + 1,
                               "IndexM:=", ModesList[n]['IndexM'],
                               "IndexN:=", ModesList[n]['IndexN'],
                               "Attenuation:=", ModesList[n]['Attenuation'],
                               "PolarizationState:=", ModesList[n]['Polarization'],
                               "AffectsRefinement:=", ModesList[n]['AffectsRefinement']])

    latticeAarray = ["NAME:LatticeAVector",
			         "Coordinate System:=", "Global",
			         "Start:=", startCoordsA,
                     "End:=", endCoordsA]
    latticeBarray = ["NAME:LatticeBVector",
			         "Coordinate System:=", "Global",
			         "Start:=", startCoordsB,
                     "End:=", endCoordsB]

    floquetportarray = ["NAME:" + portname,
                         "Faces:=", faceidlist,
                         "NumModes:=", Nmodes,
                         "RenormalizeAllTerminals:=", RenormalizeAllTerminals,
                         "DoDeembed:=", DoDeembed,
                         "DeembedDist:=", Ex(DeembedDistance).expr,
                         modesarray,
                         "ShowReporterFilter:=", ShowReporterFilter,
                         "UseScanAngles:=", UseScanAngles,
                         "Phi:=", Ex(Phi).expr+"deg",
                         "Theta:=", Ex(Theta).expr+"deg",
                         latticeAarray,
                         latticeBarray,
                         modesListArray]

    oBoundarySetupModule = get_module(oDesign, "BoundarySetup")
    oBoundarySetupModule.AssignFloquetPort(floquetportarray)
