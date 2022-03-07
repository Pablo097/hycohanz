# -*- coding: utf-8 -*-
"""
HFSS functions that use the 3D modeler. Functions in this module correspond
more or less to the functions described in the HFSS Scripting Guide,
Section "3D Modeler Editor Script Commands".

At last count there were 37 functions implemented out of 93.
"""

from __future__ import division, print_function, unicode_literals, absolute_import

import warnings

import hycohanz.conf as conf
from hycohanz.expression import Expression as Ex

warnings.simplefilter('default')

# TODO: Functions that admit list of parts or dictionaries as a way of
# generalizing for multiple inputs, could check if the input parameters are
# instead single values and also admit that, in order not to have to create
# a list for a single input value and that type of things.

# List with possible conductor materials, in order to automatically set the
# SolveInside flag to False in the corresponding functions. This list should
# be updated whenever more conductors are needed.
# Indeed, this list is also visible from outside in order to let the user
# modify it from the executable scripts.
conductors_list = ['pec', 'copper', 'silver', 'gold']

@conf.checkDefaultEditor
def get_matched_object_name(oEditor, name_filter="*"):
    """
    Returns a list of objects that match the input filter.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    name_filter : str
        Wildcard text to search.  Should contain '*' wildcard character for tuples.

    Returns
    -------
    part : list
        List of object names matched to the filter.

    """

    selections = oEditor.GetMatchedObjectName(name_filter)

    return list(selections)

@conf.checkDefaultEditor
def assign_material(oEditor, partlist, MaterialName='vacuum', SolveInside=None):
    """
    Assign a material to the specified objects. Only the MaterialName and
    SolveInside parameters of <AttributesArray> are supported.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to which the material is applied.

    Returns
    -------
    None
    """
    if SolveInside == None:
        if MaterialName in conductors_list:
            SolveInside = False
        else:
            SolveInside = True

    selectionsarray = ["NAME:Selections",
                       "Selections:=", ','.join(partlist)]

    attributesarray = ["NAME:Attributes",
                       "MaterialName:=", MaterialName,
                       "SolveInside:=", SolveInside]

    oEditor.AssignMaterial(selectionsarray, attributesarray)

@conf.checkDefaultEditor
def create_rectangle(   oEditor,
                        xs,
                        ys,
                        zs,
                        width,
                        height,
                        WhichAxis='Z',
                        Name='Rectangle1',
                        Flags='',
                        Color=(132, 132, 193),
                        Transparency=0,
                        PartCoordinateSystem='Global',
                        UDMId='',
                        MaterialName='vacuum',
                        SolveInside=None,
                        IsCovered=True):
    """
    Draw a rectangle.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    xs : float or hycohanz Expression object
    ys : float or hycohanz Expression object
    zs : float or hycohanz Expression object
        The x, y, and z coordinates of the start of the rectangle.
    width : float or hycohanz Expression object
        x-dimension of the rectangle
    height : float or hycohanz Expression object
        y-dimension of the rectangle
    WhichAxis : str
        The axis normal to the circle.  Can be 'X', 'Y', or 'Z'.
    Name : str
        The requested name of the object.  If this is not available, HFSS
        will assign a different name, which is returned by this function.
    Flags : str
        Flags associated with this object, as "NonModel". See HFSS Scripting Guide for details.
    Color : tuple of length=3
        RGB components of the circle
    Transparency : float between 0 and 1
        Fractional transparency.  0 is opaque and 1 is transparent.
    PartCoordinateSystem : str
        The name of the coordinate system in which the object is drawn.
    MaterialName : str
        Name of the material to assign to the object.
    SolveInside : bool
        Whether to mesh the interior of the object and solve for the fields
        inside.
    IsCovered : bool
        Whether the rectangle is has a surface or has only edges.

    Returns
    -------
    str
        The actual name of the created object.

    """
    if SolveInside == None:
        if MaterialName in conductors_list:
            SolveInside = False
        else:
            SolveInside = True

    RectangleParameters = [ "NAME:RectangleParameters",
                            "IsCovered:=", IsCovered,
                            "XStart:=", Ex(xs).expr,
                            "YStart:=", Ex(ys).expr,
                            "ZStart:=", Ex(zs).expr,
                            "Width:=", Ex(width).expr,
                            "Height:=", Ex(height).expr,
                            "WhichAxis:=", WhichAxis]

    Attributes = [  "NAME:Attributes",
                    "Name:=", Name,
                    "Flags:=", Flags,
                    "Color:=", "({r} {g} {b})".format(r=Color[0], g=Color[1], b=Color[2]),
                    "Transparency:=", Transparency,
                    "PartCoordinateSystem:=", PartCoordinateSystem,
                    "UDMId:=", UDMId,
                    "MaterialValue:=", '"'+MaterialName+'"',
                    "SolveInside:=", SolveInside]

    return oEditor.CreateRectangle(RectangleParameters, Attributes)

@conf.checkDefaultEditor
def create_EQbasedcurve(   oEditor,
                        xt,
                        yt,
                        zt,
                        tstart,
                        tend,
                        numpoints,
                        Version=1,
                        Name='EQcurve1',
                        Flags='',
                        Color=(132, 132, 193),
                        Transparency=0,
                        PartCoordinateSystem='Global',
                        UDMId='',
                        MaterialName='vacuum',
                        SolveInside=None):
    """
    Draw an equation based curve.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    xt : equation for the x axis, equation must be a string
    yt : equation for the y axis, equation must be a string
    zt : equation for the z axis, equation must be a string
    tstart : the start of variable _t, must be a string
    tend : then end of variable _t, must be a string
    numpoints : The number of points that the curve is made of,
        enter 0 for a continuous line, must be a string
    Version Number : ?
        Unsure what this does
    Name : str
        The requested name of the object.  If this is not available, HFSS
        will assign a different name, which is returned by this function.
    Flags : str
        Flags associated with this object.  See HFSS help for details.
    Color : tuple of length=3
        RGB components of the circle
    Transparency : float between 0 and 1
        Fractional transparency.  0 is opaque and 1 is transparent.
    PartCoordinateSystem : str
        The name of the coordinate system in which the object is drawn.
    MaterialName : str
        Name of the material to assign to the object.
    SolveInside : bool
        Whether to mesh the interior of the object and solve for the fields
        inside.

    Returns
    -------
    str
        The actual name of the created object.

    """
    if SolveInside == None:
        if MaterialName in conductors_list:
            SolveInside = False
        else:
            SolveInside = True

    EquationCurveParameters = [ "NAME:EquationBasedCurveParameters",
                            "XtFunction:=", xt,
                            "YtFunction:=", yt,
                            "ZtFunction:=", zt,
                            "tStart:=", tstart,
                            "tEnd:=", tend,
                            "NumOfPointsOnCurve:=", numpoints,
                            "Version:=", Version]

    Attributes = [  "NAME:Attributes",
                    "Name:=", Name,
                    "Flags:=", Flags,
                    "Color:=", "({r} {g} {b})".format(r=Color[0], g=Color[1], b=Color[2]),
                    "Transparency:=", Transparency,
                    "PartCoordinateSystem:=", PartCoordinateSystem,
                    "UDMId:=", UDMId,
                    "MaterialValue:=", '"'+MaterialName+'"',
                    "SolveInside:=", SolveInside]

    return oEditor.CreateEquationCurve(EquationCurveParameters, Attributes)

@conf.checkDefaultEditor
def create_circle(oEditor, xc, yc, zc, radius,
                  WhichAxis='Z',
                  NumSegments=0,
                  Name='Circle1',
                  Flags='',
                  Color=(132, 132, 193),
                  Transparency=0,
                  PartCoordinateSystem='Global',
                  MaterialName='vacuum',
                  SolveInside=None):
    """
    Create a circle primitive.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    xc : float or hycohanz Expression object
    yc : float or hycohanz Expression object
    zc : float or hycohanz Expression object
        The x, y, and z coordinates of the center of the circle.
    radius : float
        The radius of the circle.
    WhichAxis : str
        The axis normal to the circle.  Can be 'X', 'Y', or 'Z'.
    NumSegments : int
        If 0, the circle is not segmented.  Otherwise, the circle is
        segmented into NumSegments sides.
    Name : str
        The requested name of the object.  If this is not available, HFSS
        will assign a different name, which is returned by this function.
    Flags : str
        Flags associated with this object.  See HFSS help for details.
    Color : tuple of length=3
        RGB components of the circle
    Transparency : float between 0 and 1
        Fractional transparency.  0 is opaque and 1 is transparent.
    PartCoordinateSystem : str
        The name of the coordinate system in which the object is drawn.
    MaterialName : str
        Name of the material to assign to the object.
    SolveInside : bool
        Whether to mesh the interior of the object and solve for the fields
        inside.

    Returns
    -------
    str
        The actual name of the created object.
    """
    if SolveInside == None:
        if MaterialName in conductors_list:
            SolveInside = False
        else:
            SolveInside = True

    circleparams = ["NAME:CircleParameters",
                    "XCenter:=", Ex(xc).expr,
                    "YCenter:=", Ex(yc).expr,
                    "ZCenter:=", Ex(zc).expr,
                    "Radius:=", Ex(radius).expr,
                    "WhichAxis:=", str(WhichAxis),
                    "NumSegments:=", str(NumSegments)]

    attributesarray = ["NAME:Attributes",
                       "Name:=", Name,
                       "Flags:=", Flags,
                       "Color:=", "({r} {g} {b})".format(r=Color[0], g=Color[1], b=Color[2]),
                       "Transparency:=", str(Transparency),
                       "PartCoordinateSystem:=", PartCoordinateSystem,
                       "MaterialValue:=", '"'+MaterialName+'"',
                       "Solveinside:=", SolveInside]

    return oEditor.CreateCircle(circleparams, attributesarray)

@conf.checkDefaultEditor
def create_cylinder(oEditor, xc, yc, zc, radius, height,
                  WhichAxis='Z',
                  NumSides=0,
                  Name='Circle1',
                  Flags='',
                  Color=(132, 132, 193),
                  Transparency=0,
                  PartCoordinateSystem='Global',
                  MaterialName='vacuum',
                  SolveInside=None):
    """
    Create a cylinder primitive.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    xc : float or hycohanz Expression object
    yc : float or hycohanz Expression object
    zc : float or hycohanz Expression object
        The x, y, and z coordinates of the bottom center of the cylinder.
    radius : float
        The radius of the cylinder.
    height: float
        The height of the cylinder
    WhichAxis : str
        The axis normal to the cylinder.  Can be 'X', 'Y', or 'Z'.
    NumSides : int
        If 0, the cylinder is not segmented.  Otherwise, the cylinder is
        segmented into NumSides sides.
    Name : str
        The requested name of the object.  If this is not available, HFSS
        will assign a different name, which is returned by this function.
    Flags : str
        Flags associated with this object.  See HFSS help for details.
    Color : tuple of length=3
        RGB components of the cylinder
    Transparency : float between 0 and 1
        Fractional transparency.  0 is opaque and 1 is transparent.
    PartCoordinateSystem : str
        The name of the coordinate system in which the object is drawn.
    MaterialName : str
        Name of the material to assign to the object.
    SolveInside : bool
        Whether to mesh the interior of the object and solve for the fields
        inside.

    Returns
    -------
    str
        The actual name of the created object.
    """
    if SolveInside == None:
        if MaterialName in conductors_list:
            SolveInside = False
        else:
            SolveInside = True

    cylinderparams = ["NAME:CylinderParameters",
                    "XCenter:=", Ex(xc).expr,
                    "YCenter:=", Ex(yc).expr,
                    "ZCenter:=", Ex(zc).expr,
                    "Radius:=", Ex(radius).expr,
                    "Height:=", Ex(height).expr,
                    "WhichAxis:=", Ex(WhichAxis).expr,
                    "NumSides:=", Ex(NumSides).expr]

    attributesarray = ["NAME:Attributes",
                       "Name:=", Name,
                       "Flags:=", Flags,
                       "Color:=", "({r} {g} {b})".format(r=Color[0], g=Color[1], b=Color[2]),
                       "Transparency:=", str(Transparency),
                       "PartCoordinateSystem:=", PartCoordinateSystem,
                       "UDMId:=", "",
                       "MaterialValue:=", '"'+MaterialName+'"',
                       "SurfaceMaterialValue:=", '""',
                       "Solveinside:=", SolveInside,
                       "IsMaterialEditable:=", True,
                       "UseMaterialAppearance:=", False,
                       "IsLightweight:=", False]

    return oEditor.CreateCylinder(cylinderparams, attributesarray)

@conf.checkDefaultEditor
def create_sphere(oEditor, x, y, z, radius,
                  Name="Sphere1",
                  Flags="",
                  Color=(132, 132, 193),
                  Transparency=0,
                  PartCoordinateSystem="Global",
                  UDMId="",
                  MaterialName='vacuum',
                  SolveInside=None):
    """
    Create a sphere primitive.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    x : float or hycohanz Expression object
        x position in Cartesian coordinates.
    y : float or hycohanz Expression object
        y position in Cartesian coordinates.
    z : float or hycohanz Expression object
        z position in Cartesian coordinates.
    radius : float
        The sphere radius.
    Name : str
        The requested name of the object.  HFSS doesn't necessarily honor this.
    Flags : str
        Flags to attach to the object.  See the HFSS help for explanation
        of this parameter.
    Color : tuple of ints
        The RGB indices corresponding to the desired color of the object.
    Transparency : float between 0 and 1
        Fractional transparency.  0 is opaque and 1 is transparent.
    PartCoordinateSystem : str
        The name of the coordinate system in which the object is drawn.
    UDMId : str
        Unknown use.  See HFSS documentation for explanation.
    MaterialName : str
        Name of the material to assign to the object
    SolveInside : bool
        Whether to mesh the interior of the object and solve for the fields
        inside.

    Returns
    -------
    part : str
        The actual name assigned by HFSS to the part.

    """
    if SolveInside == None:
        if MaterialName in conductors_list:
            SolveInside = False
        else:
            SolveInside = True

    sphereparametersarray = ["NAME:SphereParameters",
                             "XCenter:=", Ex(x).expr,
                             "YCenter:=", Ex(y).expr,
                             "ZCenter:=", Ex(z).expr,
                             "Radius:=", Ex(radius).expr]

    attributesarray = ["NAME:Attributes",
                       "Name:=",  Name,
                       "Flags:=", Flags,
                       "Color:=", "({r} {g} {b})".format(r=Color[0], g=Color[1], b=Color[2]),
                       "Transparency:=", Transparency,
                       "PartCoordinateSystem:=", PartCoordinateSystem,
                       "UDMId:=", UDMId,
                       "MaterialValue:=", '"'+MaterialName+'"',
                       "SolveInside:=", SolveInside]

    part = oEditor.CreateSphere(sphereparametersarray, attributesarray)

    return part

@conf.checkDefaultEditor
def create_box( oEditor,
                xpos,
                ypos,
                zpos,
                xsize,
                ysize,
                zsize,
                Name='Box1',
                Flags='',
                Color=(132, 132, 193),
                Transparency=0,
                PartCoordinateSystem='Global',
                UDMId='',
                MaterialName='vacuum',
                SolveInside=None,
                IsCovered=True,
                ):
    """
    Draw a 3D box.

    Note:  This function was contributed by C. A. Donado Morcillo.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    xpos : float or hycohanz Expression object
    ypos : float or hycohanz Expression object
    zpos : float or hycohanz Expression object
        The x, y, and z coordinates of the base point of the box.
    xsize : float or hycohanz Expression object
    ysize : float or hycohanz Expression object
    zsize : float or hycohanz Expression object
        x-, y-, and z-dimensions of the box
    Name : str
        The requested name of the object.  If this is not available, HFSS
        will assign a different name, which is returned by this function.
    Flags : str
        Flags associated with this object.  See HFSS Scripting Guide for details.
    Color : tuple of length=3
        RGB components of the circle
    Transparency : float between 0 and 1
        Fractional transparency.  0 is opaque and 1 is transparent.
    PartCoordinateSystem : str
        The name of the coordinate system in which the object is drawn.
    MaterialName : str
        Name of the material to assign to the object.
    SolveInside : bool
        Whether to mesh the interior of the object and solve for the fields
        inside.
    IsCovered : bool
        Whether the rectangle is has a surface or has only edges.

    Returns
    -------
    str
        The actual name of the created object.

    """
    if SolveInside == None:
        if MaterialName in conductors_list:
            SolveInside = False
        else:
            SolveInside = True

    BoxParameters = [ "NAME:BoxParameters",
                    "XPosition:=", Ex(xpos).expr,
                    "YPosition:=", Ex(ypos).expr,
                    "ZPosition:=", Ex(zpos).expr,
                    "XSize:=", Ex(xsize).expr,
                    "YSize:=", Ex(ysize).expr,
                    "ZSize:=", Ex(zsize).expr]

    Attributes = [  "NAME:Attributes",
                    "Name:=", Name,
                    "Flags:=", Flags,
                    "Color:=", "({r} {g} {b})".format(r=Color[0], g=Color[1], b=Color[2]),
                    "Transparency:=", Transparency,
                    "PartCoordinateSystem:=", PartCoordinateSystem,
                    "UDMId:=", UDMId,
                    "MaterialValue:=", '"'+MaterialName+'"',
                    "SolveInside:=", SolveInside]

    return oEditor.CreateBox(BoxParameters, Attributes)

@conf.checkDefaultEditor
def create_polyline(oEditor, x, y, z, Name="Polyline1",
                                Flags="",
                                Color="(132 132 193)",
                                Transparency=0,
                                PartCoordinateSystem="Global",
                                UDMId="",
                                MaterialName='vacuum',
                                SolveInside=None,
                                IsPolylineCovered=True,
                                IsPolylineClosed=False,
                                XSectionBendType="Corner",
                                XSectionNumSegments="0",
                                XSectionHeight="0mm",
                                XSectionTopWidth="0mm",
                                XSectionWidth="0mm",
                                XSectionOrient="Auto",
                                XSectionType="None",
                                SegmentType="Line",
                                NoOfPoints=2):
    """
    Draw a polyline.

    Warning:  HFSS 13 crashes when you click on the last segment in the model tree.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which to perform the operation
    x : array_like
        The x locations of the polyline vertices. Can have numeric or string elements.
    y : array_like
        The y locations of the polyline vertices. Can have numeric or string elements.
    z : array_like
        The z locations of the polyline vertices. Can have numeric or string elements.
    Name : str
        Requested name of the polyline
    Flags : str
        Certain flags that can be set, as "NonModel". See HFSS Scripting Manual for details.
    Transparency : float
        Fractional transparency of the object.  0 is opaque, 1 is transparent.
    PartCoordinateSystem : str
        Coordinate system to use in constructing the object.
    UDMId : str
        TODO:  Add documentation here.
    MaterialName : str
        Name of the material to assign to the object.
    SolveInside : bool
        Whether fields are computed inside the object.
    IsPolylineCovered : bool
        Whether the polyline is covered.
    IsPolylineClosed : bool
        Whether the polyline should be considered closed.
    TODO:  finish documentation of this function.

    Returns
    -------
    polyname : str
        Actual name of the polyline


    x, y, and z are lists with numeric or string elements.

    Example Usage
    -------------
    >>> import Hyphasis as hfss
    >>> [oAnsoftApp, oDesktop] = hfss.setup_interface()
    >>> oProject = hfss.new_project(oDesktop)
    >>> oDesign = hfss.insert_design(oProject, "HFSSDesign1", "DrivenModal")
    >>> oEditor = hfss.set_active_editor(oDesign, "3D Modeler")
    >>> tri = hfss.create_polyline(oEditor, [0, 1, 0], [0, 0, 1], [0, 0, 0])
    """
    if SolveInside == None:
        if MaterialName in conductors_list:
            SolveInside = False
        else:
            SolveInside = True

    Npts = len(x)
    polylinepoints = ["NAME:PolylinePoints"]

    for n in range(0, Npts):
        if isinstance(x[n], (str, Ex)):
            xpt = Ex(x[n]).expr
        elif isinstance(x[n], (float, int)):
            xpt = str(x[n]) + "meter"
        else:
            raise TypeError('x must be of type str, int, float, or Ex')

        if isinstance(y[n], (str, Ex)):
            ypt = Ex(y[n]).expr
        elif isinstance(y[n], (float, int)):
            ypt = str(y[n]) + "meter"
        else:
            raise TypeError('y must be of type str, int, float, or Ex')

        if isinstance(z[n], (str, Ex)):
            zpt = Ex(z[n]).expr
        elif isinstance(z[n], (float, int)):
            zpt = str(z[n]) + "meter"
        else:
            raise TypeError('z must be of type str, int, float, or Ex')

        polylinepoints.append([["NAME:PLPoint",
                        "X:=", xpt,
                        "Y:=", ypt,
                        "Z:=", zpt]])

    polylinesegments = ["NAME:PolylineSegments"]
    if IsPolylineClosed == True:
        Nsegs = Npts - 1
    else:
        Nsegs = Npts - 1

    for n in range(0, Nsegs):
        polylinesegments.append(["NAME:PLSegment",
                                 "SegmentType:=", SegmentType,
                                 "StartIndex:=", n,
                                 "NoOfPoints:=", NoOfPoints])

    polylinexsection = ["NAME:PolylineXSection",
                        "XSectionType:=", XSectionType,
                        "XSectionOrient:=", XSectionOrient,
                        "XSectionWidth:=", XSectionWidth,
                        "XSectionTopWidth:=", XSectionTopWidth,
                        "XSectionHeight:=", XSectionHeight,
                        "XSectionNumSegments:=", XSectionNumSegments,
                        "XSectionBendType:=", XSectionBendType]

    polylineparams = ["NAME:PolylineParameters",
                      "IsPolylineCovered:=", IsPolylineCovered,
                      "IsPolylineClosed:=", IsPolylineClosed,
                      polylinepoints,
                      polylinesegments]#,
#                      polylinexsection]

    polylineattribs = ["NAME:Attributes",
                       "Name:=", Name,
                       "Flags:=", Flags,
                       "Color:=", Color,
                       "Transparency:=", Transparency,
                       "PartCoordinateSystem:=", PartCoordinateSystem,
                       "UDMId:=", UDMId,
                       "MaterialValue:=", '"'+MaterialName+'"',
                       "SolveInside:=",  SolveInside]

    polyname = oEditor.CreatePolyline(polylineparams, polylineattribs)

    return polyname

@conf.checkDefaultEditor
def get_selections(oEditor):
    """
    Get a list of the currently-selected objects in the design.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.

    Returns
    -------
    selectionlist : list
        List of the selectable objects in the design?
    """
    return oEditor.GetSelections()

@conf.checkDefaultEditor
def move(oEditor, partlist, x, y, z, NewPartsModelFlag="Model"):
    """
    Move specified parts.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to be moved.
    x : float
        x displacement in Cartesian coordinates.
    y : float
        y displacement in Cartesian coordinates.
    z : float
        z displacement in Cartesian coordinates.

    Returns
    -------
    None
    """
    selectionsarray = ["NAME:Selections",
                       "Selections:=", ','.join(partlist),
                       "NewPartsModelFlag:=", NewPartsModelFlag]

    moveparametersarray = ["NAME:TranslateParameters",
                           "TranslateVectorX:=", str(Ex(x).expr),
                           "TranslateVectorY:=", str(Ex(y).expr),
                           "TranslateVectorZ:=", str(Ex(z).expr)]

    oEditor.Move(selectionsarray, moveparametersarray)

@conf.checkDefaultEditor
def get_object_name(oEditor, index):
    """
    Return the object name corresponding to the zero-based creation-order index.

    Note:  This is NOT the inverse of get_object_id_by_name()!
    """
    return oEditor.GetObjectName(index)

@conf.checkDefaultEditor
def copy(oEditor, partlist):
    """
    Copy specified parts to the clipboard.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS Editor in which to perform the operation
    partlist : list of strings
        The parts to copy

    Returns
    -------
    None
    """
    selectionsarray = ["NAME:Selections",
                       "Selections:=", ','.join(partlist),
                       "NewPartsModelFlag:=", "Model"]

    oEditor.Copy(selectionsarray)

@conf.checkDefaultEditor
def get_object_id_by_name(oEditor, objname):
    """
    Return the object ID of the specified part.

    Note:  This is NOT the inverse of get_object_name()!
    """
    return oEditor.GetObjectIDByName(objname)

@conf.checkDefaultEditor
def paste(oEditor):
    """
    Paste a design in the active project from the clipboard.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS Editor in which to perform the operation

    Returns
    -------
    pastelist : list
        List of parts that are pasted
    """
    pastelist = oEditor.Paste()
    return pastelist

@conf.checkDefaultEditor
def imprint(oEditor, blanklist, toollist, KeepOriginals=False):
    """
    Imprint an object onto another object.

    Note:  This function is undocumented in the HFSS Scripting Guide.
    """
    imprintselectionsarray = [ "NAME:Selections",
                               "Blank Parts:=", ','.join(blanklist),
                               "Tool Parts:=", ','.join(toollist)]

    imprintparams = ["NAME:ImprintParameters",
                     "KeepOriginals:=", KeepOriginals]

    return oEditor.Imprint(imprintselectionsarray, imprintparams)

@conf.checkDefaultEditor
def duplicate_along_line(oEditor, partlist, x, y, z, clonesNumber,
                        CreateNewObjectsFlag=False,
                        CreateGroupsForNewObjectsFlag=False,
                        DuplicateAssignments=True):
    """
    Duplicates specified parts along a line.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to be mirrored.
    x : float or hycohanz Expression object
        x component of the sweep vector.
    y : float or hycohanz Expression object
        y component of the sweep vector.
    z : float or hycohanz Expression object
        z component of the sweep vector.
    clonesNumber : int or hycohanz Expression object
        Number of duplicated clones, including the original objects
    CreateNewObjectsFlag : boolean
        Tells whether to create new objects for each duplicated one
        or keep all the clones and original as only one object

    Returns
    -------
    selectionlist : list
        List with the original and duplicated parts names
    """
    selectionsarray = ["NAME:Selections",
                       "Selections:=", ','.join(partlist),
                       "NewPartsModelFlag:=", "Model"]

    duplicateparamsarray = ["NAME:DuplicateToAlongLineParameters",
                            "CreateNewObjects:="    , CreateNewObjectsFlag,
                            "XComponent:="      , Ex(x).expr,
                            "YComponent:="      , Ex(y).expr,
                            "ZComponent:="      , Ex(z).expr,
                            "NumClones:="       , Ex(clonesNumber).expr]

    objectName = oEditor.DuplicateAlongLine(selectionsarray, duplicateparamsarray,
        [
            "NAME:Options",
            "DuplicateAssignments:=", DuplicateAssignments
        ],
        [
            "CreateGroupsForNewObjects:=", CreateGroupsForNewObjectsFlag
        ])

    return partlist + list(objectName)

@conf.checkDefaultEditor
def duplicate_around_axis(oEditor, partlist, angle, clonesNumber, axis="Z",
                        CreateNewObjectsFlag=False,
                        CreateGroupsForNewObjectsFlag=False,
                        DuplicateAssignments=True):
    """
    Duplicates specified parts around a given axis.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to be mirrored.
    axis : str
        Rotation axis.
    angle : float or hycohanz Expression object
        Rotation angle in radians
    clonesNumber : int or hycohanz Expression object
        Number of duplicated clones, including the original objects
    CreateNewObjectsFlag : boolean
        Tells whether to create new objects for each duplicated one
        or keep all the clones and original as only one object

    Returns
    -------
    selectionlist : list
        List with the original and duplicated parts names
    """
    selectionsarray = ["NAME:Selections",
                       "Selections:=", ','.join(partlist),
                       "NewPartsModelFlag:=", "Model"]

    duplicateparamsarray = ["NAME:DuplicateAroundAxisParameters",
                            "CreateNewObjects:="    , CreateNewObjectsFlag,
                            "WhichAxis:="      , axis,
                            "AngleStr:="        , Ex(angle).expr,
                            "NumClones:="       , Ex(clonesNumber).expr]

    objectName = oEditor.DuplicateAroundAxis(selectionsarray, duplicateparamsarray,
        [
            "NAME:Options",
            "DuplicateAssignments:=", DuplicateAssignments
        ],
        [
            "CreateGroupsForNewObjects:=", CreateGroupsForNewObjectsFlag
        ])

    return partlist + list(objectName)

@conf.checkDefaultEditor
def duplicate_mirror(oEditor, partlist, base, normal,
                        CreateGroupsForNewObjectsFlag=False,
                        DuplicateAssignments=True):
    """
    Duplicate-mirror specified parts about a given base point with respect
    to a given plane.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to be duplicate-mirrored.
    base : list
        Mirror base point in Cartesian coordinates.
    normal : list
        Mirror plane normal in Cartesian coordinates.

    Returns
    -------
    selectionlist : list
        List with the original and duplicated parts names
    """
    selectionsarray = ["NAME:Selections",
                       "Selections:=", ','.join(partlist),
                       "NewPartsModelFlag:=", "Model"]

    mirrorparamsarray = ["NAME:DuplicateToMirrorParameters",
                         "DuplicateMirrorBaseX:=", Ex(base[0]).expr,
                         "DuplicateMirrorBaseY:=", Ex(base[1]).expr,
                         "DuplicateMirrorBaseZ:=", Ex(base[2]).expr,
                         "DuplicateMirrorNormalX:=", Ex(normal[0]).expr,
                         "DuplicateMirrorNormalY:=", Ex(normal[1]).expr,
                         "DuplicateMirrorNormalZ:=", Ex(normal[2]).expr]

    objectName = oEditor.DuplicateMirror(selectionsarray, mirrorparamsarray,
        [
            "NAME:Options",
            "DuplicateAssignments:=", DuplicateAssignments
        ],
        [
            "CreateGroupsForNewObjects:=", CreateGroupsForNewObjectsFlag
        ])

    return partlist + list(objectName)

@conf.checkDefaultEditor
def mirror(oEditor, partlist, base, normal):
    """
    Mirror specified parts about a given base point with respect to a given
    plane.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to be mirrored.
    base : list
        Mirror base point in Cartesian coordinates.
    normal : list
        Mirror plane normal in Cartesian coordinates.

    Returns
    -------
    None
    """
    selectionsarray = ["NAME:Selections",
                       "Selections:=", ','.join(partlist),
                       "NewPartsModelFlag:=", "Model"]

    mirrorparamsarray = ["NAME:MirrorParameters",
                         "MirrorBaseX:=", Ex(base[0]).expr,
                         "MirrorBaseY:=", Ex(base[1]).expr,
                         "MirrorBaseZ:=", Ex(base[2]).expr,
                         "MirrorNormalX:=", Ex(normal[0]).expr,
                         "MirrorNormalY:=", Ex(normal[1]).expr,
                         "MirrorNormalZ:=", Ex(normal[2]).expr]

    oEditor.Mirror(selectionsarray, mirrorparamsarray)

    return get_selections(oEditor)

@conf.checkDefaultEditor
def sweep_along_vector(oEditor, obj_name_list, x, y, z):
    """
    Sweeps the specified 1D or 2D parts along a vector.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    obj_name_list : list
        List of part name strings to be scaled.
    x : float
        x component of the sweep vector.
    y : float
        y component of the sweep vector.
    z : float
        z component of the sweep vector.

    Returns
    -------
    selectionlist : list
        List of the scaled parts names
    """
    selections = ", ".join(obj_name_list)

#    print(selections)

    oEditor.SweepAlongVector(["NAME:Selections",
                              "Selections:=", selections,
                              "NewPartsModelFlag:=", "Model"],
                             ["NAME:VectorSweepParameters",
                              "DraftAngle:=", "0deg",
                              "DraftType:=", "Round",
                              "CheckFaceFaceIntersection:=", False,
                              "SweepVectorX:=", Ex(x).expr,
                              "SweepVectorY:=", Ex(y).expr,
                              "SweepVectorZ:=", Ex(z).expr])

    return get_selections(oEditor)

@conf.checkDefaultEditor
def rotate(oEditor, partlist, axis, angle):
    """
    Rotate specified parts.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to be rotated.
    axis : str
        Rotation axis.
    angle : float
        Rotation angle in radians

    Returns
    -------
    None
    """
    selectionsarray = ["NAME:Selections",
                       "Selections:=", ','.join(partlist),
                       "NewPartsModelFlag:=", "Model"]

    rotateparametersarray = ["NAME:RotateParameters",
                             "RotateAxis:=", axis,
                             "RotateAngle:=", Ex(angle).expr]

    oEditor.Rotate(selectionsarray, rotateparametersarray)

@conf.checkDefaultEditor
def subtract(oEditor, blanklist, toollist, KeepOriginals=False):
    """
    Subtract the specified objects.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to be subtracted.
    toollist : list
        List of part name strings to subtract from partlist.
    KeepOriginals : bool
        Whether to clone the tool parts.

    Returns
    -------
    objname : str
        Name of object created by the subtract operation
    """
#    blankliststr = ""
#    for item in blanklist:
#        blankliststr += (',' + item)
#
#    toolliststr = ""
#    for item in toollist:
#        toolliststr += (',' + item)

    subtractselectionsarray = ["NAME:Selections",
                               "Blank Parts:=", ','.join(blanklist),
                               "Tool Parts:=", ','.join(toollist)]

    subtractparametersarray = ["NAME:SubtractParameters",
                               "KeepOriginals:=", KeepOriginals]

    oEditor.Subtract(subtractselectionsarray, subtractparametersarray)

    return blanklist[0]

@conf.checkDefaultEditor
def unite(oEditor, partlist, KeepOriginals=False):
    """
    Unite the specified objects.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to be united.
    KeepOriginals : bool
        Whether to keep the original parts for subsequent operations.

    Returns
    -------
    objname : str
        Name of object created by the unite operation

    Examples
    --------
    >>> import Hyphasis as hfss
    >>> [oAnsoftApp, oDesktop] = hfss.setup_interface()
    >>> oProject = hfss.new_project(oDesktop)
    >>> oDesign = hfss.insert_design(oProject, "HFSSDesign1", "DrivenModal")
    >>> oEditor = hfss.set_active_editor(oDesign, "3D Modeler")
    >>> tri1 = hfss.create_polyline(oEditor, [0, 1, 0], [0, 0, 1], [0, 0, 0])
    >>> tri2 = hfss.create_polyline(oEditor, [0, -1, 0], [0, 0, 1], [0, 0, 0])
    >>> tri3 = hfss.unite(oEditor, [tri1, tri2])
    """
#    partliststr = ""
#    for item in partlist:
#        partliststr += (',' + item)

    selectionsarray = ["NAME:Selections", "Selections:=", ','.join(partlist)]

    uniteparametersarray = ["NAME:UniteParameters", "KeepOriginals:=", KeepOriginals]

    oEditor.Unite(selectionsarray, uniteparametersarray)

    return partlist[0]

@conf.checkDefaultEditor
def scale(oEditor, partlist, x, y, z):
    """
    Scale specified parts.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to be scaled.
    x : float
        x scaling factor in Cartesian coordinates.
    y : float
        y scaling factor in Cartesian coordinates.
    z : float
        z scaling factor in Cartesian coordinates.

    Returns
    -------
    None
    """
    selections = ", ".join(partlist)
    selectionsarray = ["NAME:Selections",
                       "Selections:=", selections,
                       "NewPartsModelFlag:=", "Model"]

    scaleparametersarray = ["NAME:ScaleParameters",
                            "ScaleX:=", str(x),
                            "ScaleY:=", str(y),
                            "ScaleZ:=", str(z)]

    oEditor.Scale(selectionsarray, scaleparametersarray)

@conf.checkDefaultEditor
def get_object_name_by_faceid(oEditor, faceid):
    """
    Return the object name corresponding to the given face ID.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    faceid : int
        The face ID of the given face.

    Returns
    -------
    objname : str
        The name of the object.

    """
    return oEditor.GetObjectNameByFaceID(faceid)

@conf.checkDefaultEditor
def import_model(oEditor,
                 sourcefile,
                 HealOption=1,
                 CheckModel=False,
                 Options='-1',
                 FileType='UnRecognized',
                 MaxStitchTol=-1,
                 ImportFreeSurfaces=False):
    """
    Import a 3D model from a file.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    sourcefile : str
        Name of the 3D model file.

    Returns
    -------
    None

    Notes
    -----
    - This function is barely documented in the HFSS Scripting Guide.
    - No documentation of the optional arguments is given because their
      equivalents are not documented in the HFSS Scripting Guide.
    - There is no documented way to request a particular name for the created
      object.
    - There is no documented way to obtain the name assigned to the created
      object.

    Examples
    --------
    >>> import Hyphasis as hfss
    >>> [oAnsoftApp, oDesktop] = hfss.setup_interface()
    >>> oProject = hfss.new_project(oDesktop)
    >>> oDesign = hfss.insert_design(oProject, "HFSSDesign1", "DrivenModal")
    >>> oEditor = hfss.set_active_editor(oDesign, "3D Modeler")
    >>> hfss.import_model(oEditor, "Z:\shared\Parts\MachinescrewCap4-40_375mil\91251A108.SAT")
    >>> hfss.import_model(oEditor, "Z:\shared\Parts\MachinescrewCap4-40_375mil\91251A108.IGS")
    >>> hfss.import_model(oEditor, "Z:\shared\Parts\MachinescrewCap4-40_375mil\91251A108.STEP")

    """
    import_params_array =["NAME:NativeBodyParameters",
                          "HealOption:=", HealOption,
                          "CheckModel:=", CheckModel,
                          "Options:=", Options,
                          "FileType:=", FileType,
                          "MaxStitchTol:=", MaxStitchTol,
                          "ImportFreeSurfaces:=", ImportFreeSurfaces,
                          "SourceFile:=", sourcefile]

    oEditor.Import(import_params_array)

    return get_selections(oEditor)

@conf.checkDefaultEditor
def get_edge_by_position(oEditor, bodyname, x, y, z):
    """
    Get the edge of a given body that lies at a given position.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    bodyname : str
        Name of the body whose edge will be returned
    x : float
        x position of the edge
    y : float
        y position of the edge
    z : float
        z position of the edge

    Returns
    -------
    edgeid : int
        Id number of the edge.
    """
#    print(Ex(x).expr)
#    print(Ex(y).expr)
#    print(Ex(z).expr)
    positionparameters = ["NAME:EdgeParameters",
                          "BodyName:=", bodyname,
                          "Xposition:=", Ex(x).expr,
                          "YPosition:=", Ex(y).expr,
                          "ZPosition:=", Ex(z).expr]

    edgeid = oEditor.GetEdgeByPosition(positionparameters)

    return edgeid

@conf.checkDefaultEditor
def fillet(oEditor, partlist, edgelist, radius, vertexlist=[], setback=0):
    """
    Create fillets on the given edges.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list of strings
        List of part name strings to be filleted.
    edgelist : list of ints
        List of edge indexes to be filleted.
    radius : float
        Radius of the fillet.
    vertexlist : list
        List of vertices to chamfer
    setback : float
        The setback distance.  See the HFSS help for an explanation of this
        parameter.

    Returns
    -------
    None

    Examples
    --------
    >>> import Hyphasis as hfss
    >>> [oAnsoftApp, oDesktop] = hfss.setup_interface()
    >>> oProject = hfss.new_project(oDesktop)
    >>> oDesign = hfss.insert_design(oProject, "HFSSDesign1", "DrivenModal")
    >>> oEditor = hfss.set_active_editor(oDesign, "3D Modeler")
    >>> box1 = hfss.create_box(oEditor, 1, 1, 1, 0, 0, 0)
    >>> edge1 = hfss.get_edge_by_position(oEditor, box1, 0.5, 0, 0)
    >>> hfss.fillet(oEditor, [box1], [edge1], 0.25)

    """
    selectionsarray = ["NAME:Selections",
                       "Selections:=", ','.join(partlist),
                       "NewPartsModelFlag:=", "Model"]

    tempparams = ["NAME:FilletParameters",
                  "Edges:=", edgelist,
                  "Vertices:=", vertexlist,
                  "Radius:=",  Ex(radius).expr,
                  "Setback:=", str(setback)]

    filletparameters = ["NAME:Parameters", tempparams]

    oEditor.Fillet(selectionsarray, filletparameters)

@conf.checkDefaultEditor
def separate_body(oEditor, partlist, NewPartsModelFlag="Model"):
    """
    Separate bodies of the specified multi-lump object

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list of str
        List of object names to operate upon
    NewPartsModelFlag : str
        See the HFSS Scripting Guide.

    Returns
    -------
    newpartlist : list of str
        List of objects created by the operation.

    """
    selectionsarray = ["NAME:Selections",
                       "Selections:=", ",".join(partlist),
                       "NewPartsModelFlag:=", NewPartsModelFlag]

    oEditor.SeparateBody(selectionsarray)

    return (partlist[0],) + get_selections(oEditor)

@conf.checkDefaultEditor
def delete(oEditor, partlist):
    """
    Delete selected objects, coordinate systems, points, planes, and others.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list of str
        List of object names to delete

    Returns
    -------
    None

    """
    selectionsarray = ["NAME:Selections",
                       "Selections:=", ','.join(partlist)]

    return oEditor.Delete(selectionsarray)

@conf.checkDefaultEditor
def split(oEditor, partlist,
          NewPartsModelFlag="Model",
          SplitPlane='XY',
          WhichSide="PositiveOnly",
          SplitCrossingObjectsOnly=False,
          DeleteInvalidObjects=True):
    """
    Splits specified objects along a plane.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to be mirrored.
    SplitPlane : str
        Plane in which to split the part(s).  Allowable values are "XY", "ZX", or "YZ"
    WhichSide : str
        Side to keep.  Allowable values are "Both", "PositiveOnly", or "NegativeOnly"
    SplitCrossingObjectsOnly : bool
        If True, only splits objects that actually cross the split plane.
    DeleteInvalidObjects : bool
        If True, invalid objects generated by the operation are deleted.

    Returns
    -------
    None
    """
    selectionsarray = ["NAME:Selections",
                       "Selections:=", ",".join(partlist),
                       "NewPartsModelFlag:=", NewPartsModelFlag]
    splittoparams = ["NAME:SplitToParameters",
                     "SplitPlane:=", SplitPlane,
                     "WhichSide:=", WhichSide,
                     "SplitCrossingObjectsOnly:=", SplitCrossingObjectsOnly,
                     "DeleteInvalidObjects:=", DeleteInvalidObjects]

    return oEditor.Split(selectionsarray, splittoparams)

@conf.checkDefaultEditor
def get_face_by_position(oEditor, bodyname, x, y, z):
    """
    Get the face of a given body that lies at a given position.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    bodyname : str
        Name of the body whose face will be returned
    x : float
        x position of the face
    y : float
        y position of the face
    z : float
        z position of the face

    Returns
    -------
    faceid : int
        Id number of the face.
    """
    positionparameters = ["NAME:Parameters",
                          "BodyName:=", bodyname,
                          "Xposition:=", Ex(x).expr,
                          "YPosition:=", Ex(y).expr,
                          "ZPosition:=", Ex(z).expr]

    faceid = oEditor.GetFaceByPosition(positionparameters)

    return faceid

@conf.checkDefaultEditor
def uncover_faces(oEditor, partlist, dictoffacelists):
    """
    Uncover specified faces.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings whose faces will be uncovered.
    dictoffacelists : dict
        Dict containing part names as the keys, and lists of integer face
        ids as the values

    Returns
    -------
    None
    """
    selectionsarray = ["NAME:Selections", "Selections:=", ','.join(partlist)]

    uncoverparametersarray = ["NAME:Parameters"]
    for part in partlist:
        uncoverparametersarray += [["NAME:UncoverFacesParameters", "FacesToUncover:=", dictoffacelists[part]]]

    print('selectionsarray:  {s}'.format(s=selectionsarray))
    print('uncoverparametersarray:  {s}'.format(s=uncoverparametersarray))

    oEditor.UncoverFaces(selectionsarray, uncoverparametersarray)

@conf.checkDefaultEditor
def create_object_from_faces(oEditor, partlist, dictoffacelists, create_groups_flag=False):
    """
    Creates 2D objects from specified face(s).

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list or single string
        List of part name strings from which new objects will be created from
        their corresponding faces.
    dictoffacelists : dict
        Dict containing part names as the keys, and lists of integer face
        ids as the values. It also admits a single list or even single integer
        if only one part is being used.
    create_groups_flag : boolean
        Flag indicating if new objects should be in a group or not

    Returns
    -------
    None
    """
    is_partlist_list = isinstance(partlist, list)
    is_dictoffacelists_dict = isinstance(dictoffacelists, dict)
    if is_partlist_list and not is_dictoffacelists_dict:
        raise Exception("dictoffacelists must be a dictionary when partlist is a list.")

    selectionsarray = ["NAME:Selections", "Selections:="]
    if is_partlist_list:
        selectionsarray += [','.join(partlist)]
    else:
        selectionsarray += [partlist]
    selectionsarray += ["NewPartsModelFlag:=", "Model"]

    parametersarray = ["NAME:Parameters"]
    if is_dictoffacelists_dict:
        for part in partlist:
            facelist = dictoffacelists[part]
            if not isinstance(facelist, list):
                facelist = [facelist]
            parametersarray += [["NAME:BodyFromFaceToParameters",
                                 "FacesToDetach:=", facelist]]
    elif isinstance(dictoffacelists, list):
        parametersarray += [["NAME:BodyFromFaceToParameters",
                             "FacesToDetach:=", dictoffacelists]]
    elif isinstance(dictoffacelists, int):
        parametersarray += [["NAME:BodyFromFaceToParameters",
                             "FacesToDetach:=", [dictoffacelists]]]
    else:
        raise Exception("dictoffacelists must be either a dictionary, a list or an integer.")

    oEditor.CreateObjectFromFaces(selectionsarray, parametersarray,
                            ["CreateGroupsForNewObjects:=", create_groups_flag])

@conf.checkDefaultEditor
def connect(oEditor, partlist):
    """
    Connects specified 1-D parts to form a sheet, or specified 2-D parts to
    form a volume.:

    WARNING:  oEditor.Connect() is a very flaky operation, and the result can
    depend on the order that the parts are given to the operation among other
    seemingly random considerations.  It will very often fail on simple
    connect operations for no apparent reason.

    If you have difficulty with very strange-looking Connect() operations,
    first try to make your parts have the same number of vertices.  Second,
    try reversing the order that you give the parts to Connect().

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    partlist : list
        List of part name strings to be connected.

    Returns
    -------
    objname : str
        Name of object created by the connect operation
    """
    selectionsarray = ["NAME:Selections", "Selections:=", ','.join(partlist)]

    oEditor.Connect(selectionsarray)

    return partlist[0]

@conf.checkDefaultEditor
def rename_part(oEditor, oldname, newname):
    """
    Rename a part.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    oldname : str
        The name of the part to rename
    newname : str
        The new name to assign to the part

    Returns
    -------
    None

    """
    renameparamsarray = ["Name:Rename Data", "Old Name:=", oldname, "New Name:=", newname]

    return oEditor.RenamePart(renameparamsarray)

@conf.checkDefaultEditor
def get_face_ids(oEditor, body_name):
    """
    Get the face id list of a given body name.

    Parameters
    ----------
    oEditor : pywin32 COMObject
        The HFSS editor in which the operation will be performed.
    body_name : str
        Name of the body whose face id list will be returned

    Returns
    -------
    face_id_list : list of int
        list with face Id numbers of body_name
    """

    face_id_list = list(oEditor.GetFaceIDs(body_name))
    return list(map(int,face_id_list))
