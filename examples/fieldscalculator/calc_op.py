import hycohanz as hfss
import os.path

# Remember: with the current library version, all the oAnsoftApp, oDesktop,
# oProject, oDesign and oEditor objects can be omitted

input('Press "Enter" to connect to HFSS.>')

[oAnsoftApp, oDesktop] = hfss.setup_interface()

input('Press "Enter" to open an example project.>')

filepath = os.path.join(os.path.abspath(os.path.curdir), 'WR284.hfss')

oProject = hfss.open_project(oDesktop, filepath)

input('Press "Enter" to set the active design to HFSSDesign1.>')

oDesign = hfss.set_active_design(oProject, 'HFSSDesign1')

# input('Press "Enter" to get a handle to the Fields Reporter module.>')
#
# oFieldsReporter = hfss.get_module(oDesign, 'FieldsReporter')

input('Press "Enter" to enter the field quantity "E" in the Fields Calculator.>')

hfss.enter_qty(oDesign, 'E')

input('Open the calculator to verify that the Calculator input reads "CVc : <Ex,Ey,Ez>">')

hfss.calc_op(oDesign, 'Mag')

input('Close and reopen Calculator to verify that the "Mag" function has been applied.>')

hfss.enter_vol(oDesign, 'Polyline1')

input('Close and reopen Calculator to verify that the evaluation volume is entered.>')

hfss.calc_op(oDesign, 'Maximum')

input('Close and reopen Calculator to verify that the "Maximum" function is entered.>')

hfss.clc_eval(
    oDesign,
    'Setup1',
    'LastAdaptive',
    3.95e9,
    0.0,
    {},
    )

input('Verify that the correct answer is at the top of the stack.>')

result = hfss.get_top_entry_value(
    oDesign,
    'Setup1',
    'LastAdaptive',
    3.95e9,
    0.0,
    {},
    )

print('result: ' + str(result))

input('Press "Enter" to close the example project.>')

hfss.close_project_byhandle(oDesktop, oProject)

input('Press "Enter" to quit HFSS.>')

hfss.quit_application(oDesktop)

hfss.clean_interface()
