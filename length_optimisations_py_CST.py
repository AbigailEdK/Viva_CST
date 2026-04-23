import cst
import cst.interface
import os
import sys
import numpy as np
from cst.interface import Project
# from parameter_helper_functions import create_parameter, replace_values_with_parameters

file_dir = os.path.dirname(os.path.dirname(__file__))
project_dir = os.path.join(file_dir, "Version 1")


# -----------------------------------------------------------------------------------------------
# Parameter Functions
# -----------------------------------------------------------------------------------------------
# function to create the desired parameters in the project with values and (optional) description
def create_parameter(project: "cst.interface.Project", name: str, value: float, *, description: str = None):
    if description:
        project.schematic.StoreParameterWithDescription(name, value, description)
    else:
        project.schematic.StoreParameter(name, value)

def main():
    # connect to a running DesignEnvironment or start a new one
    de = cst.interface.DesignEnvironment.connect_to_any_or_new()

    # define the filepath to the existing project
    cst_file = os.path.join(project_dir, "Unit Resonator L Optimisation Model.cst")

    # open the project
    prj = de.open_project(cst_file)

    # work with your project

    #  Initialise model parameters 
    a = 5.6896
    b = 2.8448
    t = 0.2
    # ratios = np.arange(0.1, 1.1, 0.1)
    ratios = np.array([ 0.5, 0.7])
    L = 10
    t = 0.2

    # Initialise working variables
    widths = ratios * a
    separations = ratios * b
    lengths = np.zeros((len(ratios), len(ratios)))

    par_change = f"""
    Sub Main ()
        Dim oProject As Object
        Set oProject = GetObject(, "CSTStudio.Application").Active3D()

        If oProject Is Nothing Then
            MsgBox "Could not get the active CST project. Please make sure a project is open.", vbCritical, "Error"
            Exit Sub
        End If

        oProject.StoreParameter "L", "{L}"
        oProject.StoreParameter "t", "{t}"
        oProject.Rebuild
    End Sub
    """

    # par_change = 'Sub Main () \nStoreParameter("L", L)\nStoreParameter("t", t)\nRebuildOnParametricChange (bfullRebuild, bShowErrorMsgBox)\nEnd Sub'
    prj.schematic.execute_vba_code(par_change, timeout=None) #execute VBA script



    # # Create parameter values
    # for key, parameter in parameters.items():
    #     create_parameter(prj, key, parameter["value"], description=parameter["description"])

    # replace_values_with_parameters(prj, parameters)

    # Nested for loop to iterate over different values of w and d and save the propagation constant
    # for i, w in enumerate(widths):
    #     for j, d in enumerate(separations):
    # #         # update the parameter values in the project
    #         prj.schematic.StoreParameter("w", w)
    #         prj.schematic.StoreParameter("d", d)

            # parametric update the model to reflect the new parameter values

            #  run solver

            #  retrieve the propagation constant over frequency

            # find the propagation constant at 30 GHz

            # calculate the length for half-wavelength resonance

            # save the length in the NxN array


    # save the project
    prj.save()


if __name__ == "__main__":
    main()