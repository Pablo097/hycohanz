import hycohanz as hfss

# Remember: with the current library version, all the oAnsoftApp, oDesktop,
# oProject, oDesign and oEditor objects can be omitted

input('Press "Enter" to connect to HFSS.>')

with hfss.App() as App:

    input('Press "Enter" to create a new project.>')

    with hfss.NewProject(App.oDesktop) as P:

        input('Press "Enter" to insert a new DrivenModal design named HFSSDesign1.>')

        with hfss.InsertDesign(P.oProject, "HFSSDesign1", "DrivenModal") as D:

            input('Press "Enter" to quit HFSS.>')
