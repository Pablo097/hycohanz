import hycohanz as hfss

# Remember: with the current library version, all the oAnsoftApp, oDesktop,
# oProject, oDesign and oEditor objects can be omitted

input('Press "Enter" to connect to HFSS.>')

[oAnsoftApp, oDesktop] = hfss.setup_interface()

input('Press "Enter" to create a new project.>')

oProject = hfss.new_project(oDesktop)

input('Press "Enter" to insert a new DrivenModal design named HFSSDesign1.>')

oDesign = hfss.insert_design(oProject, "HFSSDesign1", "DrivenModal")

oBoundarySetup = hfss.get_module(oDesign, "BoundarySetup")
oMeshSetup = hfss.get_module(oDesign, "MeshSetup")
oAnalysisSetup = hfss.get_module(oDesign, "AnalysisSetup")
oOptimetrics = hfss.get_module(oDesign, "Optimetrics")
oSolutions = hfss.get_module(oDesign, "Solutions")
oFieldsReporter = hfss.get_module(oDesign, "FieldsReporter")
oRadField = hfss.get_module(oDesign, "RadField")
oUserDefinedSolutionModule = hfss.get_module(oDesign, "UserDefinedSolutionModule")

input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

del oBoundarySetup
del oMeshSetup
del oAnalysisSetup
del oOptimetrics
del oSolutions
del oFieldsReporter
del oRadField
del oUserDefinedSolutionModule

hfss.clean_interface()
# This function performs the lines corresponding to:
# del oEditor       # Not necessary in this script, actually
# del oDesign
# del oProject
# del oDesktop
# del oAnsoftApp
