# -*- coding: utf-8 -*-
"""
Functions in this module correspond more or less to the functions described
in the HFSS Scripting Guide, Section "Design Object Script Commands".

At last count there were 5 functions implemented out of 27.
"""
from __future__ import division, print_function, unicode_literals, absolute_import

import hycohanz.conf as conf

@conf.checkDefaultDesign
def get_module(oDesign, ModuleName):
    """
    Get a module handle for the given module.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design object upon which to operate.
    ModuleName : str
        Name of the module to return.  One of
            - "BoundarySetup"
            - "MeshSetup"
            - "ModelSetup"
            - "AnalysisSetup"
            - "Optimetrics"
            - "Solutions"
            - "FieldsReporter"
            - "RadField"
            - "ReportSetup"
            - "UserDefinedSolutionModule"

    Returns
    -------
    oModule : pywin32 COMObject
        Handle to the given module

    """
    oModule = oDesign.GetModule(ModuleName)

    return oModule

@conf.checkDefaultDesign
def set_active_editor(oDesign, editorname="3D Modeler"):
    """
    Set the active editor.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design upon which to operate.
    editorname : str
        Name of the editor to set as active.  As of this writing "3D Modeler"
        is the only known valid value.

    Returns
    -------
    oEditor : pywin32 COMObject
        The HFSS Editor object.

    """
    conf.oEditor = oDesign.SetActiveEditor(editorname)

    return conf.oEditor

@conf.checkDefaultDesign
def create_open_region(oDesign, frequency, Boundary="Radiation", ApplyInfiniteGP=False):
    """
    Creates an open region

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design object upon which to operate.
    frequency : float or int
        Frequency in Hz

    Returns
    -------
    None
    """
    oModule = get_module(oDesign, "ModelSetup")
    oModule.CreateOpenRegion(
		[
			"NAME:Settings",
			"OpFreq:="      , str(frequency)+"Hz",
			"Boundary:="        , Boundary,
			"ApplyInfiniteGP:=" , ApplyInfiniteGP
		])

@conf.checkDefaultDesign
def insert_infinite_sphere(oDesign,
                    name = 'Infinite Sphere 1',
                    thetaArray = [0, 180, 2],
                    phiArray = [-180, 180, 2],
                    UseCustomRadSurface = False,
                    Polarization = 'Linear',
                    UseLocalCS = False):
    """
    Creates an infinite sphere setup

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design object upon which to operate.
    thetaArray : list of int
    phiArray : list of int
        List with [start, stop, step] of theta and phi in (deg)

    Returns
    -------
    None
    """
    oModule = get_module(oDesign, "RadField")
    oModule.InsertInfiniteSphereSetup(
    	[
    		"NAME:"+name,
    		"UseCustomRadiationSurface:=", UseCustomRadSurface,
    		"CSDefinition:="	, "Theta-Phi",
    		"Polarization:="	, Polarization,
    		"ThetaStart:="		, str(thetaArray[0])+"deg",
    		"ThetaStop:="		, str(thetaArray[1])+"deg",
    		"ThetaStep:="		, str(thetaArray[2])+"deg",
    		"PhiStart:="		, str(phiArray[0])+"deg",
    		"PhiStop:="		, str(phiArray[1])+"deg",
    		"PhiStep:="		, str(phiArray[2])+"deg",
    		"UseLocalCS:="		, UseLocalCS
    	])
    return

@conf.checkDefaultDesign
def solve(oDesign,setup_name_list):
    """
    Solve Setup.

    Parameters
    ----------
    oDesktop : pywin32 COMObject
        HFSS Desktop object.

    Returns
    -------
    None

    """
    return oDesign.Solve([setup_name_list])
