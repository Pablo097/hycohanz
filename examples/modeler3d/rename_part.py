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

objname = hfss.create_polyline(
    oEditor, [1, 0, -1],
    [0, 1, 0],
    [0, 0, 0],
    Name='Triangle1',
    Color="(255 0 0)",
    )

input('Press "Enter" to rename the part as T1.>')

result = hfss.rename_part(oEditor, objname, 'T1')

print('result: ' + str(res))

input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

hfss.clean_interface()
