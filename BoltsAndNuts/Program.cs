using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.ProcessPower.PnP3dObjects;
using System.Data;
using PlantApp = Autodesk.ProcessPower.PlantInstance.PlantApplication;

//limitations:
//this was made for metric projects, for sure it needs to be modified for imperial and mixed metric projects
//for now only working with FL, WF or LUG connections
//only looking at S1 and S2 ports (not e.g. S3)
//currently the boltcircleradius is considered to belong to the bolt geometry data, but in reality it is a flange parameter, so this might need to change
//many more limitations expected due to so many special cases can exist regarding e.g. connection situation or boltset requirements
public class BoltArrayCreator
{
    [CommandMethod("CreateBoltArray", CommandFlags.UsePickSet)]
    public void CreateBoltArray()
    {
        Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
        Database db = doc.Database;
        Editor ed = doc.Editor;
        //create a special layer only for the boltset geometry, make this layer current, save the previous current layer
        SetBoltsAndNutsLayer(db, ed);

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {

            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            //select the connectors for flanged connections
            var connectors = GetFlangedConnectors(tr, ed);
            if (connectors.Count == 0)
            {
                ed.WriteMessage("\nno flanged connection found");
                return;
            }
            foreach (ObjectId connectorId in connectors)
            {
                Connector connector = (Connector)tr.GetObject(connectorId, OpenMode.ForRead);
                var subpart = GetBoltSetSubpart(connector, ed);

                if (subpart != null)
                {
                    try
                    {
                        //get all needed information from the connector and the boltset subpart
                        string boltSize = subpart.BoltSize;
                        double boltLength = subpart.BoltLength;
                        int boltCount = subpart.BoltCount;
                        string boltStandard = subpart.BoltStandard;
                        Point3d connPos = subpart.ConnPos;
                        Matrix3d connEcs = subpart.ConnEcs;

                        //get the geometry information of the bolts based on your bolts database
                        var geometryData = GetBoltGeometryData(boltStandard, boltSize);
                        if (geometryData == null)
                        {
                            ed.WriteMessage("\nno geometry data provided for this bolt standard and size");
                            continue;
                        }

                        //get the gasket thickness and the midpoint of the gasket
                        var gasketThicknessAndMidpoint = GetGasketThicknessAndMidpoint(tr, btr, ed, connectorId);
                        double gasketThickness = gasketThicknessAndMidpoint.distance;
                        Point3d gasketMidpoint = gasketThicknessAndMidpoint.midpoint;

                        //get the connected parts by clash check
                        var connectedPartsInfo = GetConnectedPartsInfo(tr, btr, ed, gasketThickness, gasketMidpoint, connectorId);

                        //create the single bolt
                        Solid3d bolt = CreateBolt(tr, btr, geometryData, connectedPartsInfo, gasketThickness, gasketMidpoint, boltSize, boltLength, ed, connPos);

                        //position the bolt and create copies in the right place
                        CreatePolarArray(tr, btr, bolt, connPos, connEcs, boltCount);
                    }
                    catch (System.Exception ex) { ed.WriteMessage("\nerror when creating bolts for connector: " + connectorId.Handle.ToString()); }
                }
                else continue;
            }
            //switch back to the saved layer
            RestoreSavedLayer(db, ed, tr);
            tr.Commit();
        }
    }

    private static string savedLayerName;

    private void SetBoltsAndNutsLayer(Database db, Editor ed)
    {

        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

            // Save the current layer
            LayerTableRecord currentLayer = tr.GetObject(db.Clayer, OpenMode.ForRead) as LayerTableRecord;
            savedLayerName = currentLayer.Name;

            // Check if current layer is "BoltsAndNuts"
            if (currentLayer.Name != "BoltsAndNuts")
            {
                if (!lt.Has("BoltsAndNuts"))
                {
                    // Create the "BoltsAndNuts" layer if it does not exist
                    lt.UpgradeOpen();
                    LayerTableRecord newLayer = new LayerTableRecord
                    {
                        Name = "BoltsAndNuts"
                    };
                    lt.Add(newLayer);
                    tr.AddNewlyCreatedDBObject(newLayer, true);
                }

                // Set "BoltsAndNuts" as the current layer
                db.Clayer = lt["BoltsAndNuts"];
            }
            tr.Commit();
        }
    }

    private void RestoreSavedLayer(Database db, Editor ed, Transaction tr)
    {
        if (string.IsNullOrEmpty(savedLayerName))
        {
            ed.WriteMessage("\nNo layer was previously saved.");
            return;
        }

        LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

        if (lt.Has(savedLayerName))
        {
            db.Clayer = lt[savedLayerName];
        }
        else
        {
            ed.WriteMessage($"The layer '{savedLayerName}' does not exist.");
        }
    }

    private ObjectIdCollection GetFlangedConnectors(Transaction tr, Editor ed)
    {
        ObjectIdCollection flangedConnectors = new ObjectIdCollection();


        // Define the filter criteria for selecting connectors
        TypedValue[] filterList = new TypedValue[]
        {
        new TypedValue((int)DxfCode.Start, "ACPPCONNECTOR"), // Select PnP3D Connectors

        };

        // Create a SelectionFilter using the filter criteria
        SelectionFilter filter = new SelectionFilter(filterList);

        PromptSelectionResult selRes = ed.SelectImplied();

        if (selRes.Status == PromptStatus.OK)
        {
            // Apply the filter to the preselected objects
            ObjectId[] SelSet = selRes.Value.GetObjectIds();

        }
        else
        {
            // Perform the selection
            selRes = ed.SelectAll(filter);
            ObjectId[] SelSet = selRes.Value.GetObjectIds();
        }



        // Check if the selection was successful
        if (selRes.Status == PromptStatus.OK)
        {
            // Iterate through the selected objects
            List<string> flangedConn = new List<string> { "Lug", "Flanged", "WaferFlanged" };
            foreach (ObjectId connectorId in selRes.Value.GetObjectIds())
            {
                if (connectorId.ObjectClass.Name.Equals("AcPpDb3dConnector"))//preselection might not be clean, so need to check again
                {
                    Connector connector = tr.GetObject(connectorId, OpenMode.ForRead) as Connector;
                    if (connector != null)
                    {
                        if (flangedConn.Contains(connector.PartSizeProperties.PropValue("JointType").ToString()))
                            flangedConnectors.Add(connectorId);
                    }
                }
            }
        }

        return flangedConnectors;
    }

    private dynamic GetBoltSetSubpart(Connector connector, Editor ed)
    {

        if (connector != null)
        {

            foreach (Autodesk.ProcessPower.PnP3dObjects.SubPart sp in connector.AllSubParts)
            {

                //btw: gasket is Autodesk.ProcessPower.PnP3dObjects.BlockSubPart

                if (sp.GetType().ToString().Equals("Autodesk.ProcessPower.PnP3dObjects.BoltSetSubPart"))
                {

                    try
                    {
                        PartSizeProperties spprops = sp.PartSizeProperties;
                        Point3d ConnPos = connector.Position;
                        Matrix3d ConnEcs = connector.Ecs;

                        //Vector3d ConnNorm = connector.XAxis;

                        /* example boltset properties:
                         subpart boltset props: PartFamilyId val: a38093dd-2004-4034-b2b9-f975a1505748
                            subpart boltset props: PartFamilyLongDesc val: Bolt set, C, 10, Stud Bolt, DIN 2501
                            subpart boltset props: PartSizeLongDesc val: Stud Bolt M16 x 150 Lg,  w/2 Hex. Nut M16, DIN 934, 2 Washer M16, DIN 125 A
                            subpart boltset props: ShortDescription val: Bolt set
                            subpart boltset props: Spec val: 10HC01
                            subpart boltset props: Size val: 100
                            subpart boltset props: WeightUnit val: KG
                            subpart boltset props: NominalDiameter val: 100
                            subpart boltset props: NominalUnit val: mm
                            subpart boltset props: ContentIsoSymbolDefinition val: TYPE=BOLT
                            subpart boltset props: Status val: New
                            subpart boltset props: AcquisitionProperties val:
                            subpart boltset props: SpecRecordId val: 12461
                            subpart boltset props: Length val: 150
                            subpart boltset props: Shop_Field val: FIELD
                            subpart boltset props: BoltSize val: M16
                            subpart boltset props: NumberInSet val: 8.00
                            subpart boltset props: BoltCompatibleStd val: DIN 976 
                         */
                        /*
                         foreach (string propertyName in spprops.PropNames)
                            {
                                object propertyValue = spprops.PropValue(propertyName);
                                ed.WriteMessage($"\nProperty Name: {propertyName}, Property Value: {propertyValue}");
                            }
                         */
                        //ed.WriteMessage("\nboltsize: " + Double.Parse(spprops.PropValue("BoltSize").ToString().Substring(1)) + " blotlenght: " + Double.Parse(spprops.PropValue("Length").ToString()) + " count: " + Convert.ToInt32(Double.Parse(spprops.PropValue("NumberInSet").ToString())) + " std: " + spprops.PropValue("BoltCompatibleStd").ToString());
                        return new
                        {
                            BoltSize = spprops.PropValue("BoltSize").ToString(),
                            BoltLength = Double.Parse(spprops.PropValue("Length").ToString()),
                            BoltCount = Convert.ToInt32(Double.Parse(spprops.PropValue("NumberInSet").ToString())),
                            BoltStandard = spprops.PropValue("BoltCompatibleStd").ToString(),
                            ConnPos,
                            ConnEcs
                        };
                    }
                    catch (System.Exception) { }

                    break;
                }
            }

        }
        return null;
    }

    private dynamic GetBoltGeometryData(string boltStandard, string boltSize)
    {
        // Fetch geometry data from table
        // This is a placeholder implementation
        // if BoltHeadHeight and BoltHeadRadius is 0, then thread rod with two nuts instead of bolt
        // if InsideHexRadius not 0, then create inside hex based on this radius and the bolt head radius

        System.Data.DataTable table = new System.Data.DataTable();

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

        //for now, because there is no real data, we use sample data
        if (boltStandard.Equals(""))
            boltStandard = "sample";

        // Display the table
        for (int i = 0; i < 2; i++)
        {
            if(i==1) boltStandard = "sample";
            foreach (DataRow row in table.Rows)
            {
                if (row["BoltStandard"].Equals(boltStandard) && row["BoltSize"].Equals(boltSize))
                {
                    return row;
                }
            }
        }
        return null;
    }

    public Solid3d CreateBolt(Transaction tr, BlockTableRecord btr, dynamic geometryData, dynamic connectedPartsInfo, double gasketThickness, Point3d gasketMidpoint, string boltSize, double boltLength, Editor ed, Point3d connPos)
    {
        // Extract geometry data
        double boltHeadHeight = geometryData["BoltHeadHeight"];
        double boltHeadRadius = geometryData["BoltHeadRadius"];
        double nutHeight = geometryData["NutHeight"];
        double nutRadius = geometryData["NutRadius"];
        double washerRadius = geometryData["WasherRadius"];
        double washerHeight = geometryData["WasherHeight"];
        double boltCircleRadius = geometryData["BoltCircleRadius"];
        double insideHexRadius = geometryData["InsideHexRadius"];
        //we get PnPClassName, FlangeThickness and S1/S2 EndType, length of both connected parts, Flange first
        double FlangeThickness1 = Double.Parse(connectedPartsInfo[0][1]);
        double FlangeThickness2 = 0;
        Point3d FlangePortPos1 = connectedPartsInfo[0][4];
        try { FlangeThickness2 = Double.Parse(connectedPartsInfo[1][1]); } catch (System.Exception) { }
        string EndType2 = connectedPartsInfo[1][2].ToString();
        double Length2 = Double.Parse(connectedPartsInfo[1][3]);
        bool isRod = false;
        if (boltHeadHeight == 0 && boltHeadRadius == 0)
        {
            boltHeadHeight = nutHeight;
            boltHeadRadius = nutRadius;
            isRod = true;
        }
        // Create the bolt head hexagon

        Solid3d boltHead = null;
        if (insideHexRadius == 0)
        {
            boltHead = CreateHexagon(tr, btr, boltHeadRadius, boltHeadHeight);
        }
        else
        {
            Solid3d boltHeadnegative = CreateHexagon(tr, btr, insideHexRadius, boltHeadHeight);
            boltHead = CreateCylinder(tr, btr, boltHeadRadius, boltHeadHeight);
            boltHead.BooleanOperation(BooleanOperationType.BoolSubtract, boltHeadnegative);
            boltHeadnegative.UpgradeOpen(); boltHeadnegative.Erase();
        }
        // Create the washer cylinder
        Solid3d washer1 = CreateCylinder(tr, btr, washerRadius, washerHeight);

        // Create the bolt body cylinder
        double boltsize = 0;
        if (boltSize.StartsWith("M")) boltsize = Double.Parse(boltSize.Replace("M", ""));
        else if (boltSize.EndsWith("\"")) boltsize = Converter.StringToDistance(boltSize.Replace("\"", ""), DistanceUnitFormat.Engineering) * 25.4;

        Solid3d boltBody = CreateCylinder(tr, btr, boltsize / 2, boltLength);

        // Create the second washer cylinder
        Solid3d washer2 = null;
        if (!EndType2.Equals("LUG")) washer2 = CreateCylinder(tr, btr, washerRadius, washerHeight);

        // Create the nut hexagon
        Solid3d nut = null;
        if (!EndType2.Equals("LUG")) nut = CreateHexagon(tr, btr, nutRadius, nutHeight);

        // Position elements
        boltHead.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, -FlangeThickness1 - boltHeadHeight - washerHeight)));
        washer1.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, -FlangeThickness1 - washerHeight)));
        if (isRod)
        {
            if (EndType2.Equals("WF"))
            {
                boltBody.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, Length2 / 2 - boltLength / 2 + gasketThickness / 2)));
            }
            else
                boltBody.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, -boltLength / 2 + gasketThickness / 2)));
        }
        else
        {
            boltBody.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, -FlangeThickness1 - washerHeight)));
        }
        // 
        if (EndType2.Equals("FL"))
        {
            washer2.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, gasketThickness + FlangeThickness2)));
            nut.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, gasketThickness + FlangeThickness2 + washerHeight)));
        }
        else if (EndType2.Equals("WF"))
        {
            washer2.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, 2 * gasketThickness + FlangeThickness1 + Length2)));
            nut.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, 2 * gasketThickness + FlangeThickness1 + washerHeight + Length2)));
        }
        //// Combine all solids
        Solid3d bolt = new Solid3d();

        bolt.BooleanOperation(BooleanOperationType.BoolUnite, boltHead);
        boltHead.UpgradeOpen(); boltHead.Erase();
        bolt.BooleanOperation(BooleanOperationType.BoolUnite, washer1);
        washer1.UpgradeOpen(); washer1.Erase();
        bolt.BooleanOperation(BooleanOperationType.BoolUnite, boltBody);
        boltBody.UpgradeOpen(); boltBody.Erase();
        if (!EndType2.Equals("LUG"))
        {
            bolt.BooleanOperation(BooleanOperationType.BoolUnite, washer2);
            washer2.UpgradeOpen(); washer2.Erase();
            bolt.BooleanOperation(BooleanOperationType.BoolUnite, nut);
            nut.UpgradeOpen(); nut.Erase();
        }
        //next two lines: corrections for automatic vs manual connection creation (gasket mirrored)
        if (EndType2.Equals("WF") && (connPos.DistanceTo(FlangePortPos1) > 0.1)) bolt.TransformBy(Matrix3d.Displacement(new Vector3d(0, 0, -Length2 - gasketThickness)));
        if (EndType2.Equals("LUG") && (connPos.DistanceTo(FlangePortPos1) > 0.1)) bolt.TransformBy(Matrix3d.Rotation(180.0 * (Math.PI / 180.0), Vector3d.YAxis, new Point3d(0, 0, gasketThickness / 2)));
        bolt.TransformBy(Matrix3d.Rotation(90.0 * (Math.PI / 180.0), Vector3d.YAxis, new Point3d(0, 0, 0)));
        bolt.TransformBy(Matrix3d.Displacement(new Vector3d(0, boltCircleRadius, 0)));
        btr.AppendEntity(bolt);
        tr.AddNewlyCreatedDBObject(bolt, true);


        return bolt;
    }

    private Solid3d CreateHexagon(Transaction tr, BlockTableRecord btr, double radius, double height)
    {
        // Create a hexagon profile
        Autodesk.AutoCAD.DatabaseServices.Polyline hexagon = new Autodesk.AutoCAD.DatabaseServices.Polyline(6);
        for (int i = 0; i < 6; i++)
        {
            double angle = i * Math.PI / 3;
            hexagon.AddVertexAt(i, new Point2d(radius * Math.Cos(angle), radius * Math.Sin(angle)), 0, 0, 0);
        }
        hexagon.Closed = true;

        // Extrude the hexagon to create a 3D solid
        DBObjectCollection curves = new DBObjectCollection();
        curves.Add(hexagon);
        // Convert the polyline to a region
        DBObjectCollection hexregion = Region.CreateFromCurves(curves);

        //btr.AppendEntity((Region)hexregion[0]);
        //tr.AddNewlyCreatedDBObject((Region)hexregion[0], true);


        // Erase the original polyline
        //polyline.UpgradeOpen();
        //polyline.Erase();

        Solid3d hexSolid = new Solid3d();
        hexSolid.Extrude((Region)hexregion[0], height, 0);

        btr.AppendEntity(hexSolid);
        tr.AddNewlyCreatedDBObject(hexSolid, true);

        return hexSolid;
    }

    private Solid3d CreateCylinder(Transaction tr, BlockTableRecord btr, double radius, double height)
    {
        // Create a circular profile
        Circle circle = new Circle(Point3d.Origin, Vector3d.ZAxis, radius);

        // Extrude the hexagon to create a 3D solid
        DBObjectCollection curves = new DBObjectCollection();
        curves.Add(circle);
        // Convert the polyline to a region
        DBObjectCollection circleregion = Region.CreateFromCurves(curves);

        //btr.AppendEntity((Region)circleregion[0]);
        //tr.AddNewlyCreatedDBObject((Region)circleregion[0], true);

        // Extrude the circle to create a cylinder
        Solid3d cylinder = new Solid3d();
        cylinder.Extrude((Region)circleregion[0], height, 0);
        //remove circle?
        btr.AppendEntity(cylinder);
        tr.AddNewlyCreatedDBObject(cylinder, true);

        return cylinder;
    }

    private Point3d GetGasketMidPoint(Connector connector)
    {
        // Implement logic to get the midpoint of the gasket
        // This is a placeholder implementation
        return Point3d.Origin;
    }

    private void CreatePolarArray(Transaction tr, BlockTableRecord btr, Solid3d bolt, Point3d centerPoint, Matrix3d orientation, int numberOfItems)
    {
        // Matrix3d rotationMatrix1 = Matrix3d.Rotation(90.0 * (Math.PI / 180.0), Vector3d.XAxis, center);
        // Apply the rotation to the original matrix
        // Matrix3d rotatedMatrix = orientation * rotationMatrix1;
        //move bolt in position
        bolt.TransformBy(orientation);

        // Parameters for the polar array
        double angleBetweenItems = 360.0 / numberOfItems; // Angle between items in degrees

        // Create the polar array
        //tmp
        //numberOfItems = 2;
        for (int i = 1; i < numberOfItems; i++)
        {
            // Calculate the rotation angle
            double rotationAngle = i * angleBetweenItems * (Math.PI / 180.0); // Convert to radians

            // Create a copy of the solid
            Solid3d newSolid = bolt.Clone() as Solid3d;

            // Transform the copy
            Matrix3d rotationMatrix = Matrix3d.Rotation(rotationAngle, orientation.CoordinateSystem3d.Xaxis, centerPoint);
            newSolid.TransformBy(rotationMatrix);
            bolt.BooleanOperation(BooleanOperationType.BoolUnite, newSolid);

            // Add the new solid to the BlockTableRecord
            btr.AppendEntity(newSolid);
            tr.AddNewlyCreatedDBObject(newSolid, true);
            newSolid.UpgradeOpen(); newSolid.Erase();
        }
        Matrix3d lastRotation = Matrix3d.Rotation(angleBetweenItems * (Math.PI / 360.0), orientation.CoordinateSystem3d.Xaxis, centerPoint);
        bolt.TransformBy(lastRotation);
    }

    private bool BoxesIntersect(Point3d amin, Point3d amax, Point3d bmin, Point3d bmax)
    {
        if (amax.X < bmin.X) return false;
        if (amin.X > bmax.X) return false;
        if (amax.Y < bmin.Y) return false;
        if (amin.Y > bmax.Y) return false;
        if (amax.Z < bmin.Z) return false;
        if (amin.Z > bmax.Z) return false;
        return true;
    }

    private dynamic GetConnectedPartsInfo(Transaction tr, BlockTableRecord btr, Editor ed, double gasketThickness, Point3d gasketMidpoint, ObjectId self)
    {
        /* example flange properties plus port properties
         Key: PnPID, Value: 1545
        Key: PnPClassName, Value: Flange
        Key: PnPStatus, Value: 0
        Key: PnPRevision, Value: 1
        Key: PnPGuid, Value: b34c0131-2942-4fe0-8d6f-3a9d35a5daa3
        Key: PnPTimestamp, Value: 638702812473073238
        Key: PartFamilyId, Value: 3146ab43-d078-4e50-a781-360e296c364e
        Key: PartFamilyLongDesc, Value: Flange C DIN 2632
        Key: CompatibleStandard, Value: DIN 2632
        Key: Manufacturer, Value:
        Key: Material, Value:
        Key: MaterialCode, Value: 1.0037
        Key: PartSizeLongDesc, Value: Flange C 100x114.3 DIN 2632-1.0037
        Key: ShortDescription, Value: Welding neck flange
        Key: Spec, Value: 10HC01
        Key: ItemCode, Value:
        Key: Size, Value: 100
        Key: DesignStd, Value:
        Key: DesignPressureFactor, Value:
        Key: Weight, Value:
        Key: WeightUnit, Value: KG
        Key: ConnectionPortCount, Value: 2
        Key: SizeRecordId, Value: 13018fa0-f72f-5797-b99c-2ea7d5543f5a
        Key: PortName, Value: S1
        Key: NominalDiameter, Value: 100
        Key: NominalUnit, Value: mm
        Key: MatchingPipeOd, Value: 114.3
        Key: EndType, Value: FL
        Key: FlangeStd, Value:
        Key: GasketStd, Value:
        Key: Facing, Value: C
        Key: FlangeThickness, Value: 20
        Key: PressureClass, Value: 10
        Key: Schedule, Value:
        Key: WallThickness, Value: 0
        Key: EngagementLength, Value: 0
        Key: LengthUnit, Value: mm
        Key: ComponentDesignation, Value:
        Key: PartCategory, Value: Flanges
        Key: ContentIsoSymbolDefinition, Value: SKEY=FLWN,TYPE=FLANGE
        Key: Position X, Value: 879.589417
        Key: Position Y, Value: 1223.042549
        Key: Position Z, Value: -451.945371
        Key: Status, Value: New
        Key: AcquisitionProperties, Value:
        Key: COG X, Value: 879.589417
        Key: COG Y, Value: 1223.042549
        Key: COG Z, Value: -477.945371
        Key: Required Spec, Value:
        Key: SpecRecordId, Value: 1947
        Key: InsulationThickness, Value:
        Key: TracingType, Value:
        Key: InsulationType, Value:
        Key: Service, Value:
        Key: TracingSpec, Value:
        Key: InsulationSpec, Value:
        Key: Tag, Value:
        Key: TieInNumber, Value:
        Key: SpoolNumber, Value:
        Key: Unit, Value:
        Key: TOP, Value: -394.79537100000005
        Key: BOP, Value: -509.095371
        Key: LineNumberTag, Value: 55555
        Key: Shop_Field, Value: SHOP
        Key: S1_SizeRecordId, Value: 13018fa0-f72f-5797-b99c-2ea7d5543f5a
        Key: S1_PortName, Value: S1
        Key: S1_NominalDiameter, Value: 100
        Key: S1_NominalUnit, Value: mm
        Key: S1_MatchingPipeOd, Value: 114.3
        Key: S1_EndType, Value: FL
        Key: S1_Facing, Value: C
        Key: S1_FlangeThickness, Value: 20
        Key: S1_PressureClass, Value: 10
        Key: S1_WallThickness, Value: 0
        Key: S1_EngagementLength, Value: 0
        Key: S1_LengthUnit, Value: mm
        Key: S2_SizeRecordId, Value: c6bd4ca3-b281-5f88-aa28-219decffb7df
        Key: S2_PortName, Value: S2
        Key: S2_NominalDiameter, Value: 100
        Key: S2_NominalUnit, Value: mm
        Key: S2_MatchingPipeOd, Value: 114.3
        Key: S2_EndType, Value: BV
        Key: S2_PressureClass, Value: 10
        Key: S2_WallThickness, Value: 0
        Key: S2_EngagementLength, Value: 0
        Key: S2_LengthUnit, Value: mm
         */

        //use a shpere in the center of the gasket for clashing
        Solid3d acSolid = new Solid3d();
        acSolid.SetDatabaseDefaults();
        acSolid.CreateSphere(gasketThickness / 1.8);//it shall collide with the neighbour parts, thats why using 1.8, not 2
        acSolid.TransformBy(Matrix3d.Displacement(Point3d.Origin.GetVectorTo(gasketMidpoint)));
        Extents3d acSolidExtents = acSolid.GeometricExtents;

        List<List<object>> twoDimReturnList = new List<List<object>>();
        List<string> possibleEndTypes = new List<string> { "FL", "WF", "LUG" };

        foreach (ObjectId id in btr)
        {
            if (id == self) continue;
            if (!PlantApp.CurrentProject.ProjectParts["Piping"].DataLinksManager.HasLinks(id)) continue;//limit clashing items to intelligent ones

            Entity entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
            if (entity != null)
            {
                Extents3d entityExtents = new Extents3d();
                try
                {
                    entityExtents = entity.GeometricExtents;
                }
                catch (Autodesk.AutoCAD.Runtime.Exception e)
                {
                    //ed.WriteMessage("\nobject with no 3d geometry: " + id.ObjectClass.Name);
                    continue;
                }

                if (BoxesIntersect(acSolidExtents.MinPoint, acSolidExtents.MaxPoint, entityExtents.MinPoint, entityExtents.MaxPoint))
                {
                    //ed.WriteMessage("collision with: " + id.Handle.ToString());
                    List<KeyValuePair<string, string>> listProp = PlantApp.CurrentProject.ProjectParts["Piping"].DataLinksManager.GetAllProperties(id, true);
                    List<KeyValuePair<string, string>> portProps = new List<KeyValuePair<string, string>>();
                    Part asset = tr.GetObject(id, OpenMode.ForRead) as Part;
                    PortCollection portCol = asset.GetPorts(PortType.All);
                    for (int j = 0; j < portCol.Count; j++)
                    {
                        SpecPort theport = asset.PortProperties(portCol[j].Name);

                        for (int i = 0; i < theport.PropCount; i++)
                        {
                            try
                            {
                                portProps.Add(new KeyValuePair<string, string>(portCol[j].Name + "_" + theport.PropNames[i], theport.PropValue(theport.PropNames[i]).ToString()));
                            }
                            catch (System.Exception) { }
                        }
                    }
                    listProp.AddRange(portProps);
                    /*foreach (var kvp in listProp)
                    {
                        ed.WriteMessage($"\nKey: {kvp.Key}, Value: {kvp.Value}");
                    }*/
                    //we need PnPClassName, FlangeThickness and S1/S2 EndType, length
                    //from S1/S2 we expect/pick one of these three: FL / LUG / WF 
                    string PnPClassName = listProp.FirstOrDefault(kvp => kvp.Key == "PnPClassName").Value;
                    //if S1_EndType is one of these, then extract the flangethickness from this port, else from S2 
                    string EndType = "";
                    string FlangeThickness = "0";
                    Point3d PortPos = new Point3d();
                    if (possibleEndTypes.Contains(EndType = listProp.FirstOrDefault(kvp => kvp.Key == "S1_EndType").Value) && portCol[0].Position.DistanceTo(gasketMidpoint) < (gasketThickness / 2) + 0.1)
                    {
                        FlangeThickness = listProp.FirstOrDefault(kvp => kvp.Key == "S1_FlangeThickness").Value;
                        PortPos = portCol[0].Position;
                    }
                    else if (possibleEndTypes.Contains(EndType = listProp.FirstOrDefault(kvp => kvp.Key == "S2_EndType").Value) && portCol[1].Position.DistanceTo(gasketMidpoint) < (gasketThickness / 2) + 0.1)
                    {
                        FlangeThickness = listProp.FirstOrDefault(kvp => kvp.Key == "S2_FlangeThickness").Value;
                        PortPos = portCol[1].Position;
                    }
                    else continue;//wrong part identified by clash test?

                    //port pos is plausible?

                    //if WF was found, then also extract the length of the object
                    string Length = "0";
                    if (EndType == "WF")
                        Length = listProp.FirstOrDefault(kvp => kvp.Key == "Length").Value;

                    //if it is a flange, insert to pos 0, if not add to end of list
                    if (PnPClassName.Equals("Flange"))
                        twoDimReturnList.Insert(0, new List<object> { PnPClassName, FlangeThickness, EndType, Length, PortPos });
                    else
                        twoDimReturnList.Add(new List<object> { PnPClassName, FlangeThickness, EndType, Length, PortPos });

                }
            }
        }
        return twoDimReturnList;
    }

    private double CalculateSmallestDimension(Point3d point1, Point3d point2)
    {
        // Calculate the dimensions of the box
        double width = Math.Abs(point2.X - point1.X);
        double height = Math.Abs(point2.Y - point1.Y);
        double depth = Math.Abs(point2.Z - point1.Z);

        // Return the smallest dimension
        return Math.Min(Math.Min(width, height), depth);
    }
    private (double distance, Point3d midpoint) GetGasketThicknessAndMidpoint(Transaction tr, BlockTableRecord btr, Editor ed, ObjectId connectorId)
    {
        double thickness = 0;
        Part asset = tr.GetObject(connectorId, OpenMode.ForRead) as Part;
        PortCollection portCol = asset.GetPorts(PortType.All);
        thickness = portCol[0].Position.DistanceTo(portCol[1].Position);

        return (thickness, MidpointCalculator.CalculateMidpoint(portCol[0].Position, portCol[1].Position));
    }
}

public class MidpointCalculator
{
    public static Point3d CalculateMidpoint(Point3d point1, Point3d point2)
    {
        double midX = (point1.X + point2.X) / 2.0;
        double midY = (point1.Y + point2.Y) / 2.0;
        double midZ = (point1.Z + point2.Z) / 2.0;

        return new Point3d(midX, midY, midZ);
    }
}

public static class Extents3dExtensions
{
    public static Extents3d ScaleBy(this Extents3d extents, double scaleFactor)
    {
        // Calculate the center of the extents
        Point3d center = new Point3d(
            (extents.MinPoint.X + extents.MaxPoint.X) / 2,
            (extents.MinPoint.Y + extents.MaxPoint.Y) / 2,
            (extents.MinPoint.Z + extents.MaxPoint.Z) / 2
        );

        // Calculate the new min and max points
        Vector3d halfSize = (extents.MaxPoint - extents.MinPoint) / 2;
        Vector3d scaledHalfSize = halfSize * scaleFactor;

        Point3d newMinPoint = center - scaledHalfSize;
        Point3d newMaxPoint = center + scaledHalfSize;

        return new Extents3d(newMinPoint, newMaxPoint);
    }
}