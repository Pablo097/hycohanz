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
                rel_permeability=1,
                cond=0,
                diel_loss_tan=0,
                mag_loss_tan=0,
                mag_saturation=0,
                lande_g=2,
                delta_h=0
                ):
    """
    Add Material.
    This function admits 3-item lists for each of the material characteristics,
    meaning that characteristic is anisotropic with each of the 3 values for x,y,z

    Parameters
    ----------
    oProject : pywin32 COMObject
        HFSS Project object.
    material_name : str
        Name of the added material.
    rel_permittivity : float or hycohanz Expression (or list of 3 if anisotropic)
    rel_permeability : float or hycohanz Expression (or list of 3 if anisotropic)
    cond : float or hycohanz Expression (or list of 3 if anisotropic)
    diel_loss_tan : float or hycohanz Expression (or list of 3 if anisotropic)
    mag_loss_tan : float or hycohanz Expression (or list of 3 if anisotropic)
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
    """
    if does_material_exist(oProject,material_name):
        msg = material_name + " already exists in the local library. No material was created"
        warnings.warn(msg)
        return msg

    propertyList = [rel_permittivity, rel_permeability, cond, diel_loss_tan, mag_loss_tan]
    propertyNames = ["permittivity", "permeability", "conductivity",
                     "dielectric_loss_tangent", "magnetic_loss_tangent"]

    mat_param = ["NAME:"+material_name]

    for i in range(5):
        if type(propertyList[i])==list and len(propertyList[i])==3:
            mat_param.append(["NAME:"+propertyNames[i],
                              "property_type:=", "AnisoProperty",
    			              "unit:=", "",
                              "component1:=", Ex(propertyList[i][0]).expr,
    			              "component2:=", Ex(propertyList[i][1]).expr,
    			              "component3:=", Ex(propertyList[i][2]).expr])
        else:
            mat_param += [propertyNames[i]+":=", Ex(propertyList[i]).expr]

    mat_param += ["saturation_mag:=", Ex(mag_saturation).expr,
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
    return True if oDefinitionManager.DoesMaterialExist(material_name) else False
