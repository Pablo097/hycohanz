import hycohanz as hfss

input('Press "Enter" to connect to HFSS.>')
with hfss.App() as App:
    input('Press "Enter" to create a new project.>')
    with hfss.NewProject(App.oDesktop) as P:
        input('Press "Enter" to insert a new design.>')
        with hfss.InsertDesign(P.oProject, "HFSSDesign1", "DrivenModal") as D:
            input('Press "Enter" to set the active editor.>')
            with hfss.SetActiveEditor(D.oDesign) as E:
                print("Switch to HFSS and draw some random objects in the editor.")
                            
                print("When done, select some of them.")
                
                input('Press "Enter" to print a list of the selected objects.>')
                
                objlist = hfss.get_selections(E.oEditor)
                
                print(objlist)
                
                input('Press "Enter" to quit HFSS.>')
