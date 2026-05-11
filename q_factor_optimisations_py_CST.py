from typing import Union, Iterator, Tuple
import cst
from cst import interface
import cst.interface
import cst.results
import os
import sys
import numpy as np
from cst.interface import Project
# from length_optimisations_py_CST import lengths


lengths = np.array([
    [5.71952979, 6.14188424, 6.6128136,  7.15492619, 7.7707952,  8.54997046, 9.49489838, 10.6317688,  10.1092016,  10.76775701],
    [5.50585423, 5.90429472, 6.3009116,  6.77194593, 7.33250782, 8.02643833, 8.90579741, 10.0332538,  11.44975112, 10.7677594 ],
    [5.46516643, 5.8042662,  6.17501313, 6.59900074, 7.11885923, 7.75311785, 8.5950654,  9.67792331, 11.14260403, 10.76779273],
    [5.44413272, 5.76338527, 6.12207898, 6.53704607, 7.02855058, 7.65481064, 8.45810301, 9.51421184, 10.97446229, 10.76773659],
    [5.44749976, 5.77146194, 6.13331225, 6.55013087, 7.04905911, 7.674027,   8.46560958, 9.50729534, 10.94976449, 10.7677831 ],
    [5.4767664,  5.81964942, 6.20779406, 6.65018382, 7.17879147, 7.81684012, 8.63167933, 9.66891824, 11.05395391, 10.76776051],
    [5.52737315, 5.92918436, 6.37607981, 6.86905397, 7.44745591, 8.13524647, 8.97425525, 9.99555953, 11.29723625, 10.7677841 ],
    [5.64792434, 6.16315927, 6.7187827,  7.31650958, 7.98848844, 8.73981007, 9.5823376,  10.55437615, 9.95519064, 10.76775798],
    [6.01496923, 6.81753241, 7.60214333, 8.38240387, 9.16139252, 9.93492268, 10.72344938, 11.5134752,  10.41025161, 10.76772333],
    [10.71620661,10.73196667,10.73726538,10.76301915,10.76465854,10.76564093,10.76640712,10.76702272,10.76743661,10.76775612]
])



# Q Factor Init
total_Q = np.zeros((lengths.shape))

file_dir = os.path.dirname(os.path.dirname(__file__))
project_dir = os.path.join(file_dir, "Version 1")

# # *************************************************************************
# # Module-level variables (can be imported and accessed from other files)
# # *************************************************************************
# # Model parameters
a = 5.6896
b = 2.8448
t = 0.1
# L = 5.0
ratios = np.arange(0.1, 1.1, 0.1)  # More granular ratios from 0.1 to 1.0 with step of 0.1

# Working variables
widths = ratios * a
separations = ratios * b

# # *************************************************************************
# # Functions
# # *************************************************************************
# # 1.
# # Function to retrieve the result TreePaths
def get_1d_Q_int_paths(mod: Union["cst.interface.Project.Model3D", "cst.results.ResultModule"]) -> Iterator[str]:
    """
    Gets all 1D Q_int results from prj and yields them.
    """
    print("Retrieving 1D Q_int result paths from the project...")
    tree_items = mod.get_tree_items()
    for tree_item in tree_items:
        tree_item_comp = tree_item.split("\\")
        # Q_int results have at least 3 "TreePathComponents"
        if len(tree_item_comp) > 2:
            if tree_item_comp[0].startswith("1D Results"):
                if tree_item_comp[1].startswith("Q Factors"):
                    if tree_item_comp[2].startswith("Total-Q") and not tree_item.endswith("Total-Q"):
                        print("tree_item_comp2] starts with Total-Q and tree_item does not end with Total-Q")
                        yield tree_item
# # 2.
# # Function to convert Result1D to numpy arrays
def result1d_to_numpy(res: "interface.Project.Result1D") -> Tuple[np.ndarray, np.ndarray]:
    """
    converts Result1D object to numpy arrays of x(float) and of y(complex) values
    """
    x = np.array(res.GetArray("x"))
    y_complex = [complex(yre, yim) for yre, yim in zip(res.GetArray("yre"), res.GetArray("yim"))]
    y = np.array(y_complex)
    return x, y

# # *************************************************************************




def main():
    # connect to a running DesignEnvironment or start a new one
    de = cst.interface.DesignEnvironment.connect_to_any_or_new()

    # define the filepath to the existing project
    cst_file = os.path.join(project_dir, "Unit Resonator Q Optimisation Model.cst")

    # open the project
    prj = de.open_project(cst_file)

    # Module-level variables are already initialized, now we'll populate lengths in the loop below


     # Nested for loop to iterate over different values of w and d and save the propagation constant
    for i, w in enumerate(widths):
        for j, d in enumerate(separations):
            # update the parameter values in the project
            length = lengths[i][j]
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
                oProject.StoreParameter "L", "{length}"
                oProject.Rebuild
            End Sub
            """
            # update the model to reflect the new parameter values
            prj.schematic.execute_vba_code(par_change, timeout=None) #execute VBA script

            # start the solver with error handling
            try:
                solver = prj.model3d.run_solver()
                # get SimulationTime
                solver_time = prj.model3d.Solver.GetTotalSimulationTime()
                print(f"Solver finished after {solver_time} seconds")
            except RuntimeError as e:
                print(f"Solver error occurred: {e}")
                print("Attempting to retrieve results anyway...")
                # Try to get solver time even if it failed
                try:
                    solver_time = prj.model3d.Solver.GetTotalSimulationTime()
                    print(f"Solver time before error: {solver_time} seconds")
                except:
                    print("Could not retrieve solver time")

            # Retrieve the Q results
            mod = prj.model3d
            item_path = next(get_1d_Q_int_paths(mod), None)  # Returns None if no items
            print(item_path)

            # Check if item_path is None
            if item_path is None:
                total_Q[i][j] = None
                print(f"No Q_int results found. Setting total_Q[{i}][{j}] to None")
            else:
                try:
                    print(f"Processing data from tree item: {item_path}")
                    result_id = prj.model3d.ResultTree.GetResultIDsFromTreeItem(item_path)
                    result = prj.model3d.ResultTree.GetResultFromTreeItem(item_path, result_id[0])
                    
                    # total_Q[i][j] = result1d_to_numpy(result)
                    total_Q[i][j] = result.GetData()

                    print(f"Total Q: {total_Q[i][j]}")
                except Exception as e:
                    print(f"Error retrieving results: {e}")
                    total_Q[i][j] = None
            
            print(f"run: {i}, {j}")

    print(total_Q)


            
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
              

    # save the project
    prj.save()


if __name__ == "__main__":
    main()

