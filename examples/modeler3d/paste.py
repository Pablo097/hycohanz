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

input('Press "Enter" to draw a sphere.>')

obj = hfss.create_sphere(oEditor, hfss.Expression("0m"),
							hfss.Expression("0m"),
							hfss.Expression("0m"),
							hfss.Expression("0.5m")/2)

input('Press "Enter" to copy the sphere to the Clipboard.>')

hfss.copy(oEditor, [obj])

input('Press "Enter" to paste the sphere from the Clipboard.>')

obj2 = hfss.paste(oEditor)[0]

input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

hfss.clean_interface()
