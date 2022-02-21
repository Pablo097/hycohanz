import hycohanz as hfss
import os.path

# Remember: with the current library version, all the oAnsoftApp, oDesktop,
# oProject, oDesign and oEditor objects can be omitted

input('Press "Enter" to connect to HFSS.>')

[oAnsoftApp, oDesktop] = hfss.setup_interface()

input('Press "Enter" to open an example project.>')

filepath = os.path.join(os.path.abspath(os.path.curdir), 'WR284.hfss')

oProject = hfss.open_project(oDesktop, filepath)

input('Press "Enter" to close the example project.>')

hfss.close_project_byhandle(oDesktop, oProject)

input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

hfss.clean_interface()
