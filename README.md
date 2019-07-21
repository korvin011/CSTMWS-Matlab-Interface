
# Introduction

This TCSTInterface class allows for communication with CST Microwave Studio from within MATLAB using Windows' COM technology. 

The main goal of this submission is to control an **existing** CST project, get and post-process the simulation results, export geometry and get various information from the project. If there is a need to **build** the geometry programmatically, there is [another good submission in Matlab File Exchange](https://se.mathworks.com/matlabcentral/fileexchange/67731-hgiddenss-cst_app) which suits better for that.

# Features

This CST-MATLAB interface features the following:

### Model control:
 - Open/close CST project, connect to the active one;
 - Store/change/read/enumerate parameters of the model with or without the model rebuild, getting parameters' expressions.
 - Copy all model parameters and their values to MATLAB workspace.
 - Enumerate/add/delete field monitors.
 - Find Run ID for the given parameter combination.

### Solving:
 - Run selected solver;
 - Preparing the CST project for evaluating the cost function in MATLAB while optimizing. It can also be used for running a custom MATLAB function as the CST simulation post-processing step ("Template Based Post-Processing"). 
 
### Retrieving results:
 - Enumerate tree items in the Navigation Tree.
 - Read 1D results from any tree item with several available filters.
 - 1D results can be queried for a specific X-coordinate (often frequency), optionally with interpolation.
 - Read S- or Z-parameters in a convenient matrix form for multi-port structures.
 - Get model parameters corresponding to each Run ID in the Result Navigator.
 - All queries for results can have a Run ID filter.
 - Read radiation field for both single-frequency and broadband field monitors.
 - Reading results for parametric sweeps done in CST. As an option, each such result can be organized in a matrix, each dimension of which corresponds to one of the swept parameters.  
 
### Exporting:
 - Export S-parameters to TOUCHSTONE file by means of CST;
 - Export current model view to an image. User can rotate the model view before exporting.
 - Export the model geometry to an STL file (triangulated objects) with surface approximation control.
 
### View control (useful for image exporting):
 - Rotate 3D view to predefined position or custom view direction (like in MATLAB "view" function).
 - Toggle wire-frame view.
 - Toggle gradient background.
 
### Getting various information:
 - About materials used in the project: name, color, transparency.
 - About geometrical objects (solids): name, component, material, color and transparency (exactly how it looks in CST), volume, mass.
 - CST license info.
 - Project units for different quantities and coefficient to convert them to SI units.
	 
In addition, a customized STL-file reader is included in order to plot geometry like they are seen in CST MWS.

One of the class methods (`ReadParametricResults`) use two custom classes (`TResultsStorage` and `TMyTable`). I apologize for not providing source code for them, but they are not yet in the state to go public :)

If other functionality is desired, please post a feature request [here](https://github.com/korvin011/CSTMWS-Matlab-Interface/issues).

# Demos / Documentation

All functionalities are well documented in the included Live Script demos.

# Bugs found?

If you encounter any errors or notice some misfunction while using the interface, please open an issue [directly in GitHub](https://github.com/korvin011/CSTMWS-Matlab-Interface/issues). 

# Acknowledgment

I would like to thank Jan Simon for his great function [`GetFullPath`](https://se.mathworks.com/matlabcentral/fileexchange/28249-getfullpath), it is very helpful for this interface.