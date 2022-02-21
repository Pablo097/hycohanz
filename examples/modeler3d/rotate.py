from __future__ import division, print_function, unicode_literals, absolute_import

from math import pi

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

input('Press "Enter" to draw a red three-vertex polyline named Triangle1.>')

obj = hfss.create_polyline(oEditor, [1, 0, -1],
                                    [0, 1, 0],
                                    [0, 0, 0],
                                    Name='Triangle1',
                                    Color="(255 0 0)")

print("We'll now rotate the triangle about the Z axis by 30 degrees.")

input('Press "Enter" to rotate the part.>')

hfss.rotate(oEditor, [obj], 'Z', 30*pi/180)

input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

hfss.clean_interface()
