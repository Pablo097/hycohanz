import hycohanz as hfss

input('Press "Enter" to connect to HFSS.>')

with hfss.App() as App:

    input('Press "Enter" to create a new project.>')
    with hfss.NewProject(App.oDesktop) as P:

        input('Press "Enter" to insert a new DrivenModal design named HFSSDesign1.>')
        with hfss.InsertDesign(P.oProject, "HFSSDesign1", "DrivenModal") as D:

            input('Press "Enter" to set the active editor to "3D Modeler" (The default and only known correct value).>')
            with hfss.SetActiveEditor(D.oDesign) as E:

                input('Press "Enter" to draw a red box named Box1.>')

                hfss.create_box(E.oEditor,  
                                1, 
                                2, 
                                3, 
                                4,
                                5,
                                6,
                                Name='Box1',
                                Color=(255, 0, 0))

                input('Press "Enter" to release HFSS.>')
