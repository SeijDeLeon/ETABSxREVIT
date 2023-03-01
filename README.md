# ETABSxREVIT
Dynamo &amp; VBA scripts to transfer 3D models between Revit and Etabs

## Dynamo
Dynamo is a graphical programming interface installed with Revit, and may be used to modify and retrieve data in Revit files. 
Full use of the .dyn file requires a licensed installation of Revit, where you can then import the script and run in the Revit model. 
Dynamo Sandbox may be used to open the file for free, shown below:
![image](https://user-images.githubusercontent.com/55897006/222258047-600ca3b4-d536-4a4b-8b27-6c913539c3bb.png)
- The Dynamo script first gathers relevant information from the active Revit model on structural members like beams, columns, and diaphragms.
- The script then sorts the members and formats the member data depending on type.
#### Precision
Before writing any data to Excel, the location (X,Y,Z) of all points must be processed to 'snap' members together correctly. 

A beam could terminate at the start of a column in a Revit building model, and this intersection is considered a 'node' for structural analysis software. It is less important for a Revit model to know that the beam and column are connected through a single node, as long as they are visually in the same location it is acceptable for drawing purposes. 

Structural software requires a much stricter definition of nodes and intersections because the software runs an analysis with matrix division, and these matrices are built on the relationships between members and nodes. To ensure that the data from Revit will create the correct nodes and intersections in ETABS, we want to reduce excessive numerical precision from the number types by rounding. This effectively pushes member endpoints to a lower degree of precision, making them 'snap' to the same (X, Y , Z) coordinates for use later in ETABS.
![image (1)](https://user-images.githubusercontent.com/55897006/222260684-3b0be0d8-e962-4b58-a05d-5901348a8fe0.png)

Finally, the data is written to a user specified Excel file using Excel plugins.
![image](https://user-images.githubusercontent.com/55897006/222262722-9a9d9f23-d657-42ba-bec9-03d196652182.png)

# VBA
A VBA script inside the Excel file uses model data (from the Dynamo/Revit output) to populate a 3D structural analysis file in ETABS with the given data. VBA can effectively control all functions of ETABS via the ETABS API, which will open up an instance of ETABS.  

A user simply needs to provide the file location and file name of the ETABS model to start the procedure. Note that the model needs to already be initialized, but can be empty. After the population is complete a timestamp is written to Excel with a summary of members and loads, or an error will show if the script was unsuccessful.
![image (2)](https://user-images.githubusercontent.com/55897006/222263526-4091bac7-57e7-4704-9489-d760c7644053.png)
