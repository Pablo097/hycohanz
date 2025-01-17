# -*- coding: utf-8 -*-
"""
Expose the HFSS Windows COM API.

Example Usage
-------------
>>> import hycohanz as hfss
>>> [oAnsoftApp, oDesktop] = hfss.setup_interface()
>>> hfss.quit_application(oDesktop)

"""

from __future__ import division, print_function, unicode_literals, absolute_import

import warnings

#import win32com.client

warnings.simplefilter('default')

# from hycohanz.conf import (oDesktop,
#                             oAnsoftApp,
#                             oProject,
#                             oDesign,
#                             oEditor
#                             )

from hycohanz.appobject import (setup_interface,
                                clean_interface)

from hycohanz.desktop import (quit_application,
                              new_project,
                              open_project,
                              close_project_byname,
                              get_active_project,
                              set_active_project,
                              close_project_byhandle,
                              close_current_project,
                              get_projects,
                              close_all_projects,
                              close_all_projects_except_current
                              )

from hycohanz.project import (get_project_name,
                              set_active_design,
                              get_active_design,
                              insert_design,
                              get_design,
                              get_top_design_list,
                              save_project,
                              save_as_project,
                              get_path,
                              rename_project
                              )

from hycohanz.property import ( constants_dict,
                                add_property,
                                set_variable,
                                get_variables,
                                get_variable_value,
                                expand_expression,
                                eval_expression
                                )

from hycohanz.design import (get_module,
                             set_active_editor,
                             create_open_region,
                             insert_infinite_sphere,
                             solve
                             )

from hycohanz.expression import Expression
from hycohanz.modeler3d import *
from hycohanz.material import ( add_material,
                                does_material_exist,
                                )

from hycohanz.analysis_setup import (insert_frequency_sweep,
                                     insert_analysis_setup,
                                     get_setups,
                                     get_sweeps)

from hycohanz.boundarysetup import (assign_perfect_e,
                                    assign_radiation,
                                    assign_perfect_h,
                                    assign_anisotropic_impedance,
                                    assign_waveport,
                                    assign_lumpedport,
                                    assign_current,
                                    create_PML,
                                    assign_primary,
                                    assign_secondary,
                                    assign_floquetport)

from hycohanz.fieldscalculator import (enter_vol,
                                       calc_op,
                                       clc_eval,
                                       enter_qty,
                                       copy_named_expr_to_stack,
                                       add_named_expr,
                                       get_top_entry_value,
                                       export_on_grid,
                                       create_field_plot)

from hycohanz.reporter import  (export_to_file,
                                get_all_report_names,
                                create_report,
                                add_traces,
                                rename_trace,
                                change_report_properties)

class App():
    """
    Context manager for HFSS App and Desktop objects.
    """
    def __enter__(self):
        """
        The win32com.client.Dispatch() function starts HFSS and assigns a
        handle to the application as oAnsoftApp.  However, oAnsoftApp
        doesn't have a method to properly deallocate itself, i.e. to shut the
        application down.  Therefore we need to do two operations:

        1. Dispatch the oAnsoftApp object using win32com.client.Dispatch()

        2. Get a oDesktop object by calling oAnsoftApp.GetAppDesktop() that
           has the oDesktop.QuitApplication() method that we can use to
           unwind the dispatch call.
        """
        print('__enter__()')
        self.oAnsoftApp, self.oDesktop = setup_interface()

        return self

    def __exit__(self, typ, val, traceback):
        """
        Destructor for the App class.  This function is empty for two related
        reasons:

        1. This class is intended to be used only with the 'with'-statement
           execution-managed environment. Plus,

        2. 'with'-statement blocks don't define a new scope, so destructors
           don't generally get called upon exit of the 'with' block.

        3. The only methods guaranteed to be run in a 'with'-block
           are __enter__() at entry, and __exit() at exit.
        """
        quit_application(self.oDesktop)
        del self.oDesktop
        del self.oAnsoftApp

        print('__exit__()')

class OpenProject():
    """
    Context manager for opening HFSS projects.
    """
    def __init__(self, oDesktop, filepath):
        self.oDesktop = oDesktop
        self.filepath = filepath


    def __enter__(self):
        self.oProject = open_project(self.oDesktop, self.filepath)

        return self

    def __exit__(self, typ, val, traceback):
        close_project_byhandle(self.oDesktop, self.oProject)

        del self.oProject
        del self.oDesktop

class NewProject():
    """
    Create an HFSS project.  See docstring for new_project for call signature.
    """
    def __init__(self, oDesktop):
        self.oDesktop = oDesktop

    def __enter__(self):
        self.oProject = new_project(self.oDesktop)

        return self

    def __exit__(self, typ, val, traceback):
        close_project_byhandle(self.oDesktop, self.oProject)

        del self.oProject
        del self.oDesktop

class SetActiveEditor():
    """
    """
    def __init__(self, oDesign):
        self.oDesign = oDesign

    def __enter__(self):
        self.oEditor = set_active_editor(self.oDesign, editorname="3D Modeler")

        return self

    def __exit__(self, typ, val, traceback):
        del self.oEditor
        del self.oDesign

class SetActiveDesign():
    """
    """
    def __init__(self, oProject, designname):
        self.oProject = oProject
        self.designname = designname

    def __enter__(self):
        self.oDesign_orig = self.oProject.GetActiveDesign()

        print(self.oDesign_orig)

        if self.oDesign_orig is not None:
            self.designname_orig = self.oDesign_orig.GetName()

        self.oDesign = set_active_design(self.oProject, self.designname)

        return self

    def __exit__(self, typ, val, traceback):
        if self.oDesign_orig is not None:
            set_active_design(self.oProject, self.designname_orig)

        del self.oDesign
        del self.oDesign_orig
        del self.oProject

class InsertDesign():
    """
    """
    def __init__(self, oProject, designname, solutiontype):
        self.oProject = oProject
        self.designname = designname
        self.solutiontype = solutiontype

    def __enter__(self):
        self.oDesign_orig = self.oProject.GetActiveDesign()

        if self.oDesign_orig is not None:
            self.designname_orig = self.oDesign_orig.GetName()

        self.oDesign = insert_design(self.oProject, self.designname, self.solutiontype)

        return self

    def __exit__(self, typ, val, traceback):
        if self.oDesign_orig is not None:
            set_active_design(self.oProject, self.designname_orig)

        del self.oDesign_orig
        del self.oDesign
        del self.oProject

class GetActiveProject():
    """
    """
    def __init__(self, oDesktop):
        self.oDesktop = oDesktop
    def __enter__(self):
        self.oProject = get_active_project(self.oDesktop)

        return self

    def __exit__(self, typ, val, traceback):
        del self.oProject
        del self.oDesktop

class GetProjects():
    """
    Get the list of open projects.  See get_projects() docstring for
    call signature.
    """
    def __init__(self, oDesktop):
        self.oDesktop = oDesktop

    def __enter__(self):
        self.oProjectlist = get_projects(self.oDesktop)

        return self

    def __exit__(self, typ, val, traceback):
        del self.oProjectlist
        del self.oDesktop

class GetModule():
    """
    """
    def __init__(self, oDesign, ModuleName):
        self.oDesign = oDesign
        self.ModuleName = ModuleName

    def __enter__(self):
        self.oModule = get_module(self.oDesign, self.ModuleName)

        return self

    def __exit__(self, typ, val, traceback):
        del self.oModule
        del self.oDesign
