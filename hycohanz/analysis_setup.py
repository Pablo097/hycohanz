# -*- coding: utf-8 -*-
"""
Functions in this module correspond more or less to the functions described 
in the HFSS Scripting Guide, Section "Analysis Setup Module Script Commands"

At last count there were 4 functions implemented out of 20.
"""
from __future__ import division, print_function, unicode_literals, absolute_import

from hycohanz.design import get_module


def insert_frequency_sweep(oDesign,
                           setupname,
                           sweepname,
                           startvalue,
                           stopvalue,
                           stepsize,
                           IsEnabled=True,
                           SetupType="LinearStep",
                           Type="Discrete",
                           SaveFields=True,
                           ExtrapToDC=False):
    """
    Insert an HFSS frequency sweep.
    
    Warning
    -------
    The API interface for this function is very susceptible to change!  It 
    currently only works for Discrete sweeps using Linear Steps.  Contributions 
    are encouraged.
    
    Parameters
    ----------
    oAnalysisSetup : pywin32 COMObject
        The HFSS Analysis Setup Module in which to insert the sweep.
    setupname : string
        The name of the setup to add
    sweepname : string
        The desired name of the sweep
    startvalue : float
        Lowest frequency in Hz.
    stopvalue : float
        Highest frequency in Hz.
    stepsize : flot
        The frequency increment in Hz.
    IsEnabled : bool
        Whether the sweep is enabled.
    SetupType : string
        The type of sweep setup to add.  One of "LinearStep", "LinearCount", 
        or "SinglePoints".  Currently only "LinearStep" is supported.
    Type : string
        The type of sweep to perform.  One of "Discrete", "Fast", or 
        "Interpolating".  Currently only "Discrete" is supported.
    Savefields : bool
        Whether to save the fields.
    ExtrapToDC : bool
        Whether extrapolation to DC is enabled.
        
    Returns
    -------
    None
    
    """
    oAnalysisSetup = oDesign.GetModule("AnalysisSetup")
    return oAnalysisSetup.InsertFrequencySweep(setupname, 
                                        ["NAME:" + sweepname, 
                                         "IsEnabled:=", IsEnabled, 
                                         "SetupType:=", SetupType, 
                                         "StartValue:=", str(startvalue) + "Hz", 
                                         "StopValue:=", str(stopvalue) + "Hz", 
                                         "StepSize:=", str(stepsize) + "Hz", 
                                         "Type:=", Type, 
                                         "SaveFields:=", SaveFields, 
                                         "ExtrapToDC:=", ExtrapToDC])


def insert_analysis_setup(oDesign, 
                          Frequency,
                          PortsOnly=True,
                          MaxDeltaS=0.02,                       
                          Name='Setup1',
                          UseMatrixConv=False,
                          MaximumPasses=20,
                          MinimumPasses=2,
                          MinimumConvergedPasses=2,
                          PercentRefinement=30,
                          IsEnabled=True,
                          BasisOrder=2,
                          UseIterativeSolver=False,
                          DoLambdaRefine=True,
                          DoMaterialLambda=True,
                          SetLambdaTarget=True,
                          Target=0.6667,
                          UseMaxTetIncrease=False,
                          PortAccuracy=2,
                          UseABCOnPort=False,
                          SetPortMinMaxTri=False,
                          EnableSolverDomains=False,
                          ThermalFeedback=False,
                          hNoAdditionalRefinementOnImport=False):
    """
    Insert an HFSS analysis setup.
    """
    oAnalysisSetup = get_module(oDesign, "AnalysisSetup")
    oAnalysisSetup.InsertSetup( "HfssDriven", 
                               ["NAME:" + Name, 
                                "Frequency:=", str(Frequency) +"Hz", 
                                "PortsOnly:=", PortsOnly, 
                                "MaxDeltaS:=", MaxDeltaS, 
                                "UseMatrixConv:=", UseMatrixConv, 
                                "MaximumPasses:=", MaximumPasses, 
                                "MinimumPasses:=", MinimumPasses, 
                                "MinimumConvergedPasses:=", MinimumConvergedPasses, 
                                "PercentRefinement:=", PercentRefinement, 
                                "IsEnabled:=", IsEnabled, 
                                "BasisOrder:=", BasisOrder, 
                                "UseIterativeSolver:=", UseIterativeSolver, 
                                "DoLambdaRefine:=", DoLambdaRefine, 
                                "DoMaterialLambda:=", DoMaterialLambda, 
                                "SetLambdaTarget:=", SetLambdaTarget, 
                                "Target:=", Target, 
                                "UseMaxTetIncrease:=", UseMaxTetIncrease, 
                                "PortAccuracy:=", PortAccuracy, 
                                "UseABCOnPort:=", UseABCOnPort, 
                                "SetPortMinMaxTri:=", SetPortMinMaxTri, 
                                "EnableSolverDomains:=", EnableSolverDomains, 
                                "ThermalFeedback:=", ThermalFeedback, 
                                "NoAdditionalRefinementOnImport:=", NoAdditionalRefinementOnImport])
                                
    return Name


def get_setups(design):
    """
    Gets the names of analysis setups in a design.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design object upon which to operate.

    Returns
    -------
    reportnames : list of str
        A list with the names of all the analysis setups in the Design
    """
    module = get_module(design, "AnalysisSetup")
    setup_list = list(module.GetSetups())
    return map(str, setup_list)


def get_sweeps(design, setup_name):
    """
    Gets the names of all sweeps in a given analysis setup.

    Parameters
    ----------
    oDesign : pywin32 COMObject
        The HFSS design object upon which to operate.
    SweepName : str
        Name of HFSS sweep to use, for example "LastAdaptive"
    """
    module = get_module(design, "AnalysisSetup")
    sweep_list = list(module.GetSweeps(setup_name))
    return map(str, sweep_list)
