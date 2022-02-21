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

input('Press "Enter" to draw a red rectangle named Rectangle1.>')

objname = hfss.create_rectangle(
    oEditor,
    1,
    2,
    3,
    4,
    5,
    Name='Rectangle1',
    Color=(255, 0, 0))

input('Press "Enter" to sweep the rectangle.')

hfss.sweep_along_vector(oEditor, [objname], 0, 0, 1)

input('Press "Enter" to quit HFSS.>')

# hfss.quit_application(oDesktop)

hfss.clean_interface()
