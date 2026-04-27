from typing import Union, Iterator, Tuple

import cst
from cst import interface
import cst.interface
import cst.results
import os
import sys
import numpy as np
from cst.interface import Project

file_dir = os.path.dirname(os.path.dirname(__file__))
project_dir = os.path.join(file_dir, "Version 1")

# *************************************************************************
# Module-level variables (can be imported and accessed from other files)
# *************************************************************************
# Model parameters
a = 5.6896
b = 2.8448
t = 0.2
L = 5.0
# ratios = np.array([0.3, 0.7])
ratios = np.arange(0.1, 1.1, 0.1)  # More granular ratios from 0.1 to 1.0 with step of 0.1

# Working variables
widths = ratios * a
separations = ratios * b
lengths = np.zeros((len(ratios), len(ratios)))  # This will be populated in main()

# *************************************************************************
# Functions
# *************************************************************************
# 1.
# Function to retrieve the result TreePaths
def get_1d_gamma_paths(mod: Union["cst.interface.Project.Model3D", "cst.results.ResultModule"]) -> Iterator[str]:
    """
    Gets all 1D Gamma results from prj and yields them.
    """
    tree_items = mod.get_tree_items()
    for tree_item in tree_items:
        tree_item_comp = tree_item.split("\\")
        # Gamma results have at least 3 "TreePathComponents"
        if len(tree_item_comp) > 2:
            if tree_item_comp[0].startswith("1D Results"):
                if tree_item_comp[1].startswith("Port Information"):
                    # if tree_item_comp[2].startswith("Gamma") and not tree_item.endswith("Gamma"):
                    if tree_item_comp[2].startswith("Gamma") and not tree_item.endswith("Gamma"):
                        yield tree_item
# 2.
# Function to convert Result1D to numpy arrays
def result1d_to_numpy(res: "interface.Project.Result1D") -> Tuple[np.ndarray, np.ndarray]:
    """
    converts Result1D object to numpy arrays of x(float) and of y(complex) values
    """
    x = np.array(res.GetArray("x"))
    y_complex = [complex(yre, yim) for yre, yim in zip(res.GetArray("yre"), res.GetArray("yim"))]
    y = np.array(y_complex)
    return x, y

# *************************************************************************




def main():
    # connect to a running DesignEnvironment or start a new one
    de = cst.interface.DesignEnvironment.connect_to_any_or_new()

    # define the filepath to the existing project
    cst_file = os.path.join(project_dir, "Unit Resonator L Optimisation Model.cst")

    # open the project
    prj = de.open_project(cst_file)

    # Module-level variables are already initialized, now we'll populate lengths in the loop below


     # Nested for loop to iterate over different values of w and d and save the propagation constant
    for i, w in enumerate(widths):
        for j, d in enumerate(separations):
            # update the parameter values in the project
            par_change = f"""
            Sub Main ()
                Dim oProject As Object
                Set oProject = GetObject(, "CSTStudio.Application").Active3D()

                If oProject Is Nothing Then
                    MsgBox "Could not get the active CST project. Please make sure a project is open.", vbCritical, "Error"
                    Exit Sub
                End If

                oProject.StoreParameter "w", "{w}"
                oProject.StoreParameter "d", "{d}"
                oProject.Rebuild
            End Sub
            """
            # update the model to reflect the new parameter values
            prj.schematic.execute_vba_code(par_change, timeout=None) #execute VBA script

            # start the solver
            solver = prj.model3d.run_solver()

            # get SimulationTime
            solver_time = prj.model3d.Solver.GetTotalSimulationTime()
            print(f"Solver finished after {solver_time} seconds")

            # Retrieve the propagation constant
            mod = prj.model3d
            item_path = next(get_1d_gamma_paths(mod), None)  # Returns None if no items


            
            print(f"Processing data from tree item: {item_path}")
            result_id = prj.model3d.ResultTree.GetResultIDsFromTreeItem(item_path)
            result = prj.model3d.ResultTree.GetResultFromTreeItem(item_path, result_id[0])
            
            frequencies, beta = result1d_to_numpy(result)

            # Isolate the imaginary part of the propagation constant (beta) and convert to real values
            beta = np.imag(beta)

            # Find the index of the frequency closest to 30 GHz
            target_frequency = 30.0
            frequency_index = np.argmin(np.abs(np.array(frequencies) - target_frequency))
            
            propagation_constant_at_30ghz = beta[frequency_index]

            print(f"Propagation constant at {frequencies[frequency_index]:.4f} GHz: {propagation_constant_at_30ghz}")

            # Calculate the length of the resonator for a half wavelength at 30 GHz
            # Store in lengths array
            # beta*L = pi for half wavelength resonance
            lengths[i, j] = (np.pi / propagation_constant_at_30ghz)*10**3 # convert m to mm (beta defined with m)
            print(f"Calculated resonator length for half wavelength at {frequencies[frequency_index]:.2f} GHz: {lengths[i, j]:.4f} mm")

            # for item_path in first_path:
            #     print(f"Processing data from tree item: {item_path}")
            #     result_id = prj.model3d.ResultTree.GetResultIDsFromTreeItem(item_path)
            #     result = prj.model3d.ResultTree.GetResultFromTreeItem(item_path, result_id[0])
                
            #     frequencies, beta = result1d_to_numpy(result)

            #     # Isolate the imaginary part of the propagation constant (beta) and convert to real values
            #     beta = np.imag(beta)

            #     # Find the index of the frequency closest to 30 GHz
            #     target_frequency = 30.0
            #     frequency_index = np.argmin(np.abs(np.array(frequencies) - target_frequency))
                
            #     propagation_constant_at_30ghz = beta[frequency_index]

            #     print(f"Propagation constant at {frequencies[frequency_index]:.4f} GHz: {propagation_constant_at_30ghz}")

            #     # Calculate the length of the resonator for a half wavelength at 30 GHz
            #     # Store in lengths array
            #     # beta*L = pi for half wavelength resonance
            #     lengths[i, j] = np.pi / propagation_constant_at_30ghz
            #     print(f"Calculated resonator length for half wavelength at {frequencies[frequency_index]:.2f} GHz: {lengths[i, j]:.4f} mm")
            
    print(lengths)       

    # save the project
    prj.save()


if __name__ == "__main__":
    main()