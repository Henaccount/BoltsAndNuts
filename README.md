# BoltsAndNuts
Code Sample: In Plant 3D create physical bolts. Use at own risk!

<pre>
Limitations:
this is sample code, use at own risk
this was made for metric projects, for sure it needs to be modified for imperial and mixed metric projects
for now only working with FL, WF or LUG connections
only looking at S1 and S2 ports (not e.g. S3)
currently the boltcircleradius is considered to belong to the bolt geometry data, but in reality it is a flange parameter, so this might need to change
many more limitations expected due to so many special cases can exist regarding e.g. connection situation or boltset requirements

Usage:
compile with MS Visual Studio or similar. Load resulting dll with "netload" command. Execute the code with Plant 3D file opened and the "CreateBoltArray" command from the command line.
This command will create bolts for all boltsets in the file. You can select single connectors before executing the command, this will only create bolts for the selection.
Bolts will be created on a new layer called "BoltsAndNuts". Every command call will create a new set of bolts, regardless if bolts already exist or not. 
You can easily select all bolts in a file by using "Quickselect" by layer.

Configuration:
You need to provide the bolt geometries yourself, the length will be taken from the calculated boltset.
A bolt will be selected by BoltCompatibleStandard property of the boltset and the size of the bolt.
This is how to provide the data (unless you change the code to read from an Excel sheet or similar):

        // Define columns
        table.Columns.Add("BoltStandard", typeof(string));
        table.Columns.Add("BoltSize", typeof(string));
        table.Columns.Add("BoltHeadHeight", typeof(double));
        table.Columns.Add("BoltHeadRadius", typeof(double));
        table.Columns.Add("NutHeight", typeof(double));
        table.Columns.Add("NutRadius", typeof(double));
        table.Columns.Add("WasherRadius", typeof(double));
        table.Columns.Add("WasherHeight", typeof(double));
        table.Columns.Add("BoltCircleRadius", typeof(double));
        table.Columns.Add("InsideHexRadius", typeof(double));

        // Add rows
        // example for rod:
        // table.Rows.Add("default", "M16", 0, 0, 13, 13.4, 13, 2, 90, 0);
        // example for hex bolt:
        // table.Rows.Add("default", "M16", 13, 13.4, 13, 13.4, 15, 2, 90, 0);
        // example for inside hex bolt:
        // table.Rows.Add("default", "M16", 13, 13.4, 13, 13.4, 15, 2, 90, 8);
        // imperial input (code might not be fully capable for imperial input..)
        // table.Rows.Add("default", "3/4\"", 13, 13.4, 13, 13.4, 15, 2, 90, 0);

        table.Rows.Add("sample", "M10", 6.5, 10, 8, 10, 10, 2, 50, 0);
        table.Rows.Add("sample", "M12", 7.5, 12, 10, 12, 12, 2.5, 65, 0);
        table.Rows.Add("sample", "M16", 10, 16, 13, 16, 17, 3, 90, 0);
        table.Rows.Add("sample", "M20", 12.5, 20, 16, 20, 21, 4, 110, 0);
        table.Rows.Add("sample", "M24", 15, 24, 19, 24, 25, 5, 130, 0);
        table.Rows.Add("sample", "M27", 17, 27, 21, 27, 28, 5.5, 150, 0);
        table.Rows.Add("sample", "M30", 18.7, 30, 24, 30, 31, 6, 170, 0);
        table.Rows.Add("sample", "M36", 22.5, 36, 29, 36, 37, 7, 210, 0);

        // if BoltHeadHeight and BoltHeadRadius is 0, then thread rod with two nuts instead of bolt
        // if InsideHexRadius not 0, then inside hex will be created based on this radius and the bolt head radius
  
  </pre>


