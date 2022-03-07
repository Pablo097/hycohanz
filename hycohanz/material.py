# -*- coding: utf-8 -*-
"""
Functions in this module correspond more or less to the functions described
in the HFSS Scripting Guide (v 2013.11), Section "Material Script Commands".

At last count there were 2 functions implemented out of 5.
"""
from __future__ import division, print_function, unicode_literals, absolute_import

import warnings

import hycohanz.conf as conf
from hycohanz.expression import Expression as Ex
from hycohanz.desktop import get_active_project

warnings.simplefilter('default')

@conf.checkDefaultProject
def add_material(oProject,
                material_name,
                rel_permittivity=1,
                rel_permeability =1,
                cond=0,
                diel_loss_tan=0,
                mag_loss_tan=0,
                mag_saturation=0,
                lande_g=2,
                delta_h=0
                ):
    """
    Add Material.

    Parameters
    ----------
    oProject : pywin32 COMObject
        HFSS Project object.
    material_name : str
        Name of the added material.
    rel_permittivity : float or hycohanz Expression
    rel_permeability : float or hycohanz Expression
    cond : float or hycohanz Expression
    diel_loss_tan : float or hycohanz Expression
    mag_loss_tan : float or hycohanz Expression
    mag_saturation : float or hycohanz Expression
    lande_g : float or hycohanz Expression
    delta_h : float or hycohanz Expression
        The relative permittivity, relative permeability, electric
        conductivity, dielectric loss tangent, magnetic loss tangent,
        magnetic saturation, Lande G factor, and delta_h associated with
        the added material.

    Returns
    -------
    None

    #### This function could be improved by admitting 3-item lists for
    #### each of the material characteristics, assuming these would mean
    #### that characteristic is anisotropic with each of the 3 values for x,y,z
    """
    if does_material_exist(oProject,material_name):
        msg = material_name + " already exists in the local library. No material was created"
        warnings.warn(msg)
        return msg
    else:
        mat_param = ["NAME:"+material_name,
                    "permittivity:=", Ex(rel_permittivity).expr,
                    "permeability:=", Ex(rel_permeability).expr,
                    "conductivity:=", Ex(cond).expr,
                    "dielectric_loss_tangent:=", Ex(diel_loss_tan).expr,
                    "magnetic_loss_tangent:=", Ex(mag_loss_tan).expr,
                    "saturation_mag:=", Ex(mag_saturation).expr,
                    "lande_g_factor:=", Ex(lande_g).expr,
                    "delta_H:=", Ex(delta_h).expr]
        oDefinitionManager = oProject.GetDefinitionManager()
        return oDefinitionManager.AddMaterial(mat_param)

@conf.checkDefaultProject
def does_material_exist(oProject,material_name):
    """
    Check if material exists.

    Parameters
    ----------
    oProject : pywin32 COMObject
        HFSS Project object.

    Returns
    -------
    Bool

    Examples
    --------
    >>> import Hyphasis as hfss
    >>>

    """
    oDefinitionManager = oProject.GetDefinitionManager()
    return oDefinitionManager.DoesMaterialExist(material_name)
