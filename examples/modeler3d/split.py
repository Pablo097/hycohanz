from __future__ import division, print_function, unicode_literals, absolute_import

import hycohanz as hfss

# Remember: with the current library version, all the oAnsoftApp, oDesktop,
# oProject, oDesign and oEditor objects can be omitted

input('Press "Enter" to connect to HFSS.>')

[oAnsoftApp, oDesktop] = hfss.setup_interface()

input('Press "Enter" to create a new project.>')

oProject = hfss.new_project(oDesktop)

input('Press "Enter" to insert a new DrivenModal design named HFSSDesign1.>')

oDesign = hfss.insert_design(oProject, "HFSSDesign1", "DrivenModal")

input('Press "Enter" to set the active editor to "3D Modeler" (The default and only known correct value).>')

oEditor = hfss.set_active_editor(oDesign)

input('Press "Enter" to insert some circle properties into the design.>')

hfss.add_property(oDesign, "xcenter", hfss.Expression("1m"))
hfss.add_property(oDesign, "ycenter", hfss.Expression("2m"))
hfss.add_property(oDesign, "zcenter", hfss.Expression("0m"))
hfss.add_property(oDesign, "diam", hfss.Expression("6m"))

input('Press "Enter" to draw a sphere using the properties.>')

objname = hfss.create_sphere(
    oEditor,
    hfss.Expression("xcenter"),
    hfss.Expression("ycenter"),
    hfss.Expression("zcenter"),
    hfss.Expression("diam")/2,
    )

input('Press "Enter" to split the sphere into two hemispheres.>')

hfss.split(oEditor, [objname], SplitPlane='XY', WhichSide='Both')

input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

hfss.clean_interface()
