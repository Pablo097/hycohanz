import hycohanz as hfss

# Remember: with the current library version, all the oAnsoftApp, oDesktop,
# oProject, oDesign and oEditor objects can be omitted

input('Press "Enter" to connect to HFSS.>')

[oAnsoftApp, oDesktop] = hfss.setup_interface()

input('Press "Enter" to create a new project.>')

oProject = hfss.new_project(oDesktop)

input('Press "Enter" to get the name of the created project.>')

projectname = hfss.get_project_name(oProject)

input('Press "Enter" to close the project.>')

hfss.close_project_byname(oDesktop, projectname)

input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

hfss.clean_interface()
