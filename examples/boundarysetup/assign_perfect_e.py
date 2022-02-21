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
hfss.add_property(oDesign, "zcenter", hfss.Expression("3m"))
hfss.add_property(oDesign, "diam", hfss.Expression("1m"))

input('Press "Enter" to draw a circle using the properties.>')

hfss.create_circle(oEditor, hfss.Expression("xcenter"),
				     hfss.Expression("ycenter"),
				     hfss.Expression("zcenter"),
				     hfss.Expression("diam")/2)

input('Press "Enter" to assign a PerfectE boundary condition on the circle.>')

# By manual inspection of the model, we can determine that the face where we
# want to insert the boundary condition is Face_10 of Circle1, corresponding to
# face index 10.  The proper way to do this is to call GetFaceByPosition()
# (either directly or through a hycohanz wrapper function) which returns this
# index.
hfss.assign_perfect_e(oDesign, "PerfectE1", [10])

input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

hfss.clean_interface()
