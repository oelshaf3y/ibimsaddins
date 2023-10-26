using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Line = Autodesk.Revit.DB.Line;
using Parameter = Autodesk.Revit.DB.Parameter;

namespace IBIMSGen.Hangers
{
    [Transaction(TransactionMode.Manual)]
    public class HangersTrial : IExternalCommand
    {
        UIDocument uidoc;
        Document doc;
        HangersFM UI;
        FilteredElementCollector linksFEC, levelsFEC;
        FamilySymbol ductHanger, pipeHanger20, pipeHanger2, pipeHanger200;
        List<string> LinksNames, levelsNames, worksetnames;
        Document LinkDoc;
        List<Level> levels;
        List<List<string>> AllWorksetNames;
        List<List<List<string>>> AllWorksetsDIMS;
        IList<Element> MechanicalEquipment, ducts, pipes, cables, floors, ductfits;
        IList<Reference> mechRefs, linkedRefs;
        List<Face> floorFacesUp, floorFacesDown;
        List<double> waterSupplyDiams, waterSupplySpaces, chilledWaterDiams, chilledWaterSpaces, drainageDiams, drainageSpaces, fireDiams, fireSpaces;
        List<double> HangDias, floorElevations;
        List<Workset> worksets;
        List<WorksetId> worksetIDs;
        Options options;
        RevitLinkInstance RLI;
        List<DuctHanger> ductHangers;
        List<PipeHanger> pipeHangers;
        List<TrayHanger> trayHangers;
        double meanFloorHeight, ductOffset, negLength;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet element)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            UI = new HangersFM();
            options = new Options();
            options.ComputeReferences = true;
            options.View = doc.ActiveView;

            ductHangers = new List<DuctHanger>();
            pipeHangers = new List<PipeHanger>();
            trayHangers = new List<TrayHanger>();
            StringBuilder sb = new StringBuilder();
            mechRefs = new List<Reference>();
            linkedRefs = new List<Reference>();
            ducts = new List<Element>();
            pipes = new List<Element>();
            cables = new List<Element>();
            floors = new List<Element>();
            floorElevations = new List<double>();
            floorFacesUp = new List<Face>();
            floorFacesDown = new List<Face>();

            ductOffset = 100 / 304.8;
            negLength = 100 / 304.80;
            HangDias = new List<double>() { 17, 22, 27, 34, 42, 52, 65, 67, 77, 82, 92, 102, 112, 127, 152, 162, 202, 227, 252, 317, 352, 402 };

            MechanicalEquipment = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).OfClass(typeof(FamilySymbol)).ToList();
            pipeHanger20 = MechanicalEquipment.Cast<FamilySymbol>().Where(x => x.FamilyName.Equals("02- PIPE HANGER ( 20 - 200 )"))?.FirstOrDefault() ?? null;
            pipeHanger200 = MechanicalEquipment.Cast<FamilySymbol>().Where(x => x.FamilyName.Equals("01- PIPE HANGER ( +200 mm )"))?.FirstOrDefault() ?? null;
            pipeHanger2 = MechanicalEquipment.Cast<FamilySymbol>().Where(x => x.FamilyName.Equals("Pipes Hanger 2"))?.FirstOrDefault() ?? null;
            ductHanger = MechanicalEquipment.Cast<FamilySymbol>().Where(x => x.FamilyName.Equals("The Lower Bridge Duct Hanger"))?.FirstOrDefault() ?? null;

            StringBuilder nullfams = new StringBuilder();
            if (pipeHanger20 == null) { nullfams.AppendLine("02- PIPE HANGER ( 20 - 200 )"); }
            if (pipeHanger200 == null) { nullfams.AppendLine("01- PIPE HANGER ( +200 mm )"); }
            if (pipeHanger2 == null) { nullfams.AppendLine("Pipes Hanger 2"); }
            if (ductHanger == null) { nullfams.AppendLine("The Lower Bridge Duct Hanger"); }
            if (ductHanger == null || pipeHanger20 == null || pipeHanger200 == null || pipeHanger2 == null)
            {
                TaskDialog.Show("Error", "Please Load Supports Family.\n" + nullfams.ToString());
                //return Result.Failed;
            }

            linksFEC = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance));
            LinksNames = linksFEC.Cast<RevitLinkInstance>()
                .Select(x => ((RevitLinkType)doc.GetElement(x.GetTypeId())))
                .Where(x => x.GetLinkedFileStatus() == LinkedFileStatus.Loaded)
                .Select(x => x.Name).ToList();
            if (LinksNames.Count == 0)
            {
                TaskDialog.Show("Error", "There are no Loaded Linked Revit detected in Project.");
                return Result.Failed;
            }
            UI.Linkes = LinksNames;
            levelsFEC = new FilteredElementCollector(doc).OfClass(typeof(Level));
            levels = levelsFEC.Cast<Level>().OrderBy(x => x.Elevation).ToList();
            levelsNames = levels.Select(x => x.Name).ToList();
            double prevHeight = double.MaxValue;
            List<double> floorHeights = new List<double>();
            foreach (double height in levels.Select(x => x.Elevation).ToList())
            {

                if (height < prevHeight)
                {
                    prevHeight = height;
                }
                else
                {
                    floorHeights.Add((height - prevHeight));
                    prevHeight = height;
                }
            }
            meanFloorHeight = floorHeights.Average();

            worksets = new FilteredWorksetCollector(doc).Where(x => x.Kind == WorksetKind.UserWorkset).ToList();
            worksetnames = worksets.Select(x => x.Name).ToList();
            worksetIDs = worksets.Select(x => x.Id).ToList();

            if (worksetnames.Count == 0)
            {
                TaskDialog.Show("Error", "Document has no UserWorksets.");
                return Result.Failed;
            }
            UI.worksetnames = worksetnames;
            UI.Levels = levelsNames;
            UI.ShowDialog();
            if (UI.canc)
            {
                return Result.Cancelled;
            }



            AllWorksetNames = UI.AllworksetsNames;
            AllWorksetsDIMS = UI.AllworksetsDIMS;

            //HACK | modified >>> needs refactoring
            waterSupplyDiams = AllWorksetsDIMS[1][0].Select(x => Convert.ToDouble(x)).Where(x => x != 0).ToList();
            waterSupplySpaces = AllWorksetsDIMS[1][1].Select(x => Convert.ToDouble(x) / 304.8).Where(x => x != 0).ToList();
            chilledWaterDiams = AllWorksetsDIMS[2][0].Select(x => Convert.ToDouble(x)).Where(x => x != 0).ToList();
            chilledWaterSpaces = AllWorksetsDIMS[2][1].Select(x => Convert.ToDouble(x) / 304.8).Where(x => x != 0).ToList();
            drainageDiams = AllWorksetsDIMS[3][0].Select(x => Convert.ToDouble(x)).Where(x => x != 0).ToList();
            drainageSpaces = AllWorksetsDIMS[3][1].Select(x => Convert.ToDouble(x) / 304.8).Where(x => x != 0).ToList();
            fireDiams = AllWorksetsDIMS[4][0].Select(x => Convert.ToDouble(x)).Where(x => x != 0).ToList();
            fireSpaces = AllWorksetsDIMS[4][1].Select(x => Convert.ToDouble(x) / 304.8).Where(x => x != 0).ToList();


            RLI = linksFEC.Cast<RevitLinkInstance>()
                .Where(x => ((RevitLinkType)doc.GetElement(x.GetTypeId())).Name == LinksNames[UI.lnk]).First();
            LinkDoc = RLI.GetLinkDocument();




            try
            {
                mechRefs = uidoc.Selection.PickObjects(ObjectType.Element, new SelectionFilterPDC(), "Select Pipes / Ducts / CableTrays.");

                //TODO | modify getSystemRank

                foreach (Reference reference in mechRefs)
                {
                    Element e = doc.GetElement(reference);
                    if (e.Category.Name == "Ducts")
                    {
                        WorksetId wsid = e.WorksetId;
                        string ws = worksetnames[worksetIDs.IndexOf(wsid)];
                        if (GetSystemRank(ws) != 0) continue;

                        Curve ductCurve = ((LocationCurve)e.Location).Curve;
                        double s1 = Math.Round(e.LookupParameter("Start Middle Elevation").AsDouble(), 6);
                        double s2 = Math.Round(e.LookupParameter("End Middle Elevation").AsDouble(), 6);
                        if (ductCurve.Length >= negLength && s1 == s2) //perfectly horizontal 
                        {
                            ducts.Add(e);
                        }

                    }
                    else if (e.Category.Name == "Pipes")
                    {
                        WorksetId wsid = e.WorksetId;
                        string ws = worksetnames[worksetIDs.IndexOf(wsid)];
                        if (GetSystemRank(ws) == -1 || GetSystemRank(ws) == 0) continue;

                        Curve pipeCurve = ((LocationCurve)e.Location).Curve;
                        if (pipeCurve.Length >= negLength && Math.Abs(Math.Round(((Line)pipeCurve).Direction.Normalize().Z, 3)) != 1)
                        {
                            pipes.Add(e);
                        }
                    }
                    else if (e is CableTray)
                    {
                        Curve c = ((LocationCurve)e.Location).Curve;
                        if (c.Length >= negLength && Math.Abs(Math.Round(((Line)c).Direction.Normalize().Z, 3)) == 0)
                        {
                            cables.Add(e);
                        }
                    }
                }
            }
            catch
            {
                return Result.Cancelled;
            }
            if (UI.selc) // allows me to select region
            {
                try
                {
                    linkedRefs = uidoc.Selection.PickObjects(ObjectType.LinkedElement, "Select Linked Elements.");

                    foreach (Reference reff in linkedRefs)//if by selection
                    {
                        RevitLinkInstance rli = doc.GetElement(reff.ElementId) as RevitLinkInstance;
                        if (rli != null)
                        {
                            if (RLI.Name.Trim() == rli.Name.Trim())
                            {
                                Document linkDocument = rli.GetLinkDocument();
                                Element ele = linkDocument.GetElement(reff.LinkedElementId);
                                if (ele is Floor)
                                {
                                    floors.Add(ele);
                                }
                            }
                        }
                    }
                }
                catch
                {
                    return Result.Cancelled;
                }
            }
            else // by levels
            {
                Level from = levels[UI.frin];
                Level to = levels[UI.toin];
                FilteredElementCollector floorLinkedFEC = new FilteredElementCollector(LinkDoc).OfClass(typeof(Floor));
                foreach (Floor floor in floorLinkedFEC)
                {
                    double elevationAtBot = 0;
                    double ElevationAtTop = 0;
                    Parameter atBotParam = floor.LookupParameter("Elevation at Bottom");
                    Parameter atTopParam = floor.LookupParameter("Elevation at Top");
                    if (atBotParam != null && atTopParam != null)
                    {
                        elevationAtBot = atBotParam.AsDouble();
                        ElevationAtTop = atTopParam.AsDouble();
                        if (elevationAtBot == 0 && ElevationAtTop == 0)
                        {
                            double levelElevation = 0;
                            Parameter levelParam = floor.LookupParameter("Level");
                            if (levelParam != null)
                            {
                                string levelName = levelParam.AsValueString();
                                foreach (Level level in levels)
                                {
                                    if (level.Name.Trim() == levelName.Trim())
                                    {
                                        levelElevation = level.Elevation;
                                        break;
                                    }
                                }
                                double heightOffset = floor.LookupParameter("Height Offset From Level").AsDouble();
                                elevationAtBot = levelElevation + heightOffset;
                                if (elevationAtBot <= to.Elevation && elevationAtBot >= from.Elevation)
                                {
                                    floors.Add(floor);
                                }
                            }
                        }
                        else
                        {
                            if (elevationAtBot <= to.Elevation && elevationAtBot >= from.Elevation)
                            {
                                floors.Add(floor);
                            }
                        }
                    }
                }
            }

            if (ducts.Count == 0 && pipes.Count == 0 && cables.Count == 0)
            {
                TaskDialog.Show("Wrong Selection", "There are no Ducts or Pipes or Cabletrays selected.");
                return Result.Failed;
            }

            if (floors.Count == 0)
            {
                TaskDialog.Show("Error", "Linked Revit has no Floors to Host Hangers." + Environment.NewLine + "Make sure that Linked Revit is Structural discipline and has Floors.");
                return Result.Failed;
            }

            //<getting region of the quad tree>
            double offset = ducts.OrderBy(x => ((LocationCurve)x.Location).Curve.Length).Select(x => ((LocationCurve)x.Location).Curve.Length).Max() * 0.6;
            double minfx = ducts.Select(x => Math.Min(((LocationCurve)x.Location).Curve.GetEndPoint(0).X, ((LocationCurve)x.Location).Curve.GetEndPoint(1).X)).Min() - offset;
            double minfy = ducts.Select(x => Math.Min(((LocationCurve)x.Location).Curve.GetEndPoint(0).Y, ((LocationCurve)x.Location).Curve.GetEndPoint(1).Y)).Min() - offset;
            double minfz = ducts.Select(x => Math.Min(((LocationCurve)x.Location).Curve.GetEndPoint(0).Z, ((LocationCurve)x.Location).Curve.GetEndPoint(1).Z)).Min() - offset;
            double maxfx = ducts.Select(x => Math.Max(((LocationCurve)x.Location).Curve.GetEndPoint(0).X, ((LocationCurve)x.Location).Curve.GetEndPoint(1).X)).Max() + offset;
            double maxfy = ducts.Select(x => Math.Max(((LocationCurve)x.Location).Curve.GetEndPoint(0).Y, ((LocationCurve)x.Location).Curve.GetEndPoint(1).Y)).Max() + offset;
            double maxfz = ducts.Select(x => Math.Max(((LocationCurve)x.Location).Curve.GetEndPoint(0).Z, ((LocationCurve)x.Location).Curve.GetEndPoint(1).Z)).Max() + offset;
            //<getting region of the quad tree>

            Boundary selectionBoundary = new Boundary(minfx, maxfy, maxfx, minfy, maxfz, minfz);
            int count = 0;
            //List<ElementId> ids = new List<ElementId>();
            ductfits = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DuctFitting)
                .OfClass(typeof(FamilyInstance))
                .Where(x => selectionBoundary.contains(x))
                .ToList();
            double minx = ductfits.Select(x => ((LocationPoint)x.Location).Point.X).Min() - offset;
            double miny = ductfits.Select(x => ((LocationPoint)x.Location).Point.Y).Min() - offset;
            double minz = ductfits.Select(x => ((LocationPoint)x.Location).Point.Z).Min() - offset;
            double maxx = ductfits.Select(x => ((LocationPoint)x.Location).Point.X).Max() + offset;
            double maxy = ductfits.Select(x => ((LocationPoint)x.Location).Point.Y).Max() + offset;
            double maxz = ductfits.Select(x => ((LocationPoint)x.Location).Point.Z).Max() + offset;
            QuadTree ductTree = new QuadTree(minx, maxy, maxx, miny, maxz, minz);
            QuadTree fitsTree = new QuadTree(minx, maxy, maxx, miny, maxz, minz);
            td("before adding ducts");
            foreach (Element duct in ducts)
            {
                if (ductTree.insert(duct)) count++;

            }
            td("ducts= " + count.ToString() + " of " + ducts.Count.ToString());
            count = 0;
            foreach (Element fitting in ductfits)
            {
                if (fitsTree.insert(fitting)) count++;
                //sb.AppendLine("fitting added: " + getLocation(fitting));
            }
            td("fits= " + count.ToString() + " of " + ductfits.Count.ToString());


            Transaction tr = new Transaction(doc, "Draw");
            tr.Start();
            foreach (Element ele in floors)
            {
                double floorLeft = double.MaxValue;
                double floorBottom = double.MaxValue;
                double floorRight = double.MinValue;
                double floorTop = double.MinValue;
                double floorUp = double.MinValue;
                double floorDown = double.MaxValue;
                Solid solid = getSolid(ele);
                if (solid != null && solid.Volume != 0)
                {
                    double elevationAtBot = 0;
                    double elevationAtTop = 0;
                    Parameter atBotParam = ele.LookupParameter("Elevation at Bottom");
                    Parameter atTopParam = ele.LookupParameter("Elevation at Top");
                    if (atBotParam != null && atTopParam != null)
                    {
                        elevationAtBot = atBotParam.AsDouble();
                        elevationAtTop = atTopParam.AsDouble();
                        if (elevationAtBot == 0 && elevationAtTop == 0)
                        {
                            double levelElevation = 0;
                            Parameter levelParam = ele.LookupParameter("Level");
                            if (levelParam != null)
                            {
                                string levelName = levelParam.AsValueString();
                                foreach (Level level in levels)
                                {
                                    if (level.Name.Trim() == levelName.Trim())
                                    {
                                        levelElevation = level.Elevation;
                                        break;
                                    }
                                }
                                double heightOffset = ele.LookupParameter("Height Offset From Level").AsDouble();
                                elevationAtBot = levelElevation + heightOffset;
                            }
                        }
                        Face face = solid.Faces.get_Item(0);

                        foreach (Face f in solid.Faces)
                        {
                            foreach (EdgeArray edgeArray in f.EdgeLoops)
                            {

                                foreach (Edge edge in edgeArray)
                                {
                                    Curve c = edge.AsCurve();
                                    XYZ p1 = c.GetEndPoint(0);
                                    XYZ p2 = c.GetEndPoint(1);
                                    if (Math.Min(p1.X, p2.X) < floorLeft) floorLeft = Math.Min(p1.X, p2.X);
                                    if (Math.Max(p1.X, p2.X) > floorRight) floorRight = Math.Max(p1.X, p2.X);
                                    if (Math.Min(p1.Y, p2.Y) < floorBottom) floorBottom = Math.Min(p1.Y, p2.Y);
                                    if (Math.Max(p1.Y, p2.Y) > floorTop) floorTop = Math.Max(p1.Y, p2.Y);
                                    if (Math.Max(p1.Z, p2.Z) > floorUp) floorUp = Math.Max(p1.Z, p2.Z);
                                    if (Math.Min(p1.Z, p2.Z) < floorDown) floorDown = Math.Min(p1.Z, p2.Z);
                                }

                            }
                        }
                        floorRight += offset;
                        floorLeft -= offset;
                        floorTop += offset;
                        floorBottom -= offset;
                        floorUp +=  meanFloorHeight;
                        floorDown -= meanFloorHeight;
                        //Boundary floorRange = new Boundary(double.MinValue, double.MaxValue, double.MaxValue, double.MinValue, double.MaxValue, double.MinValue);
                        Boundary floorRange = new Boundary(floorLeft, floorTop, floorRight, floorBottom, floorUp, floorDown);
                        List<Element> ductsInRange = ductTree.query(floorRange);
                        List<Element> fitsInRange = fitsTree.query(floorRange);

                        if (ductsInRange.Count > 0) sb.AppendLine("ducts in range " + ductsInRange.Count.ToString());
                        if (fitsInRange.Count > 0) sb.AppendLine("fits in range " + fitsInRange.Count.ToString());

                        processDucts(ductsInRange, fitsInRange, solid.Faces.get_Item(0));
                        floorFacesUp.Add(solid.Faces.get_Item(1)); //upperface
                        floorFacesDown.Add(solid.Faces.get_Item(0)); //lowerface
                        floorElevations.Add(elevationAtBot);
                        //XYZ southWestDown = new XYZ(floorLeft, floorBottom, floorDown);
                        //XYZ southWestUp = new XYZ(floorLeft, floorBottom, floorUp);
                        //XYZ northWestDown = new XYZ(floorLeft, floorTop, floorDown);
                        //XYZ northWestUp = new XYZ(floorLeft, floorTop, floorUp);
                        //XYZ northEastUp = new XYZ(floorRight, floorTop, floorUp);
                        //XYZ northEastDown = new XYZ(floorRight, floorTop, floorDown);
                        //XYZ southEastUp = new XYZ(floorRight, floorBottom, floorUp);
                        //XYZ southEastDown = new XYZ(floorRight, floorBottom, floorDown);
                        //List<GeometryObject> lines = new List<GeometryObject>();

                        //lines.Add(Line.CreateBound(southWestDown, southEastDown));
                        //lines.Add(Line.CreateBound(southWestDown, northWestDown));
                        //lines.Add(Line.CreateBound(southWestDown, southWestUp));
                        //lines.Add(Line.CreateBound(northEastDown, southEastDown));
                        //lines.Add(Line.CreateBound(northEastDown, northWestDown));
                        //lines.Add(Line.CreateBound(northEastDown, northEastUp));
                        //lines.Add(Line.CreateBound(southWestUp, southEastUp));
                        //lines.Add(Line.CreateBound(southWestUp, northWestUp));
                        //lines.Add(Line.CreateBound(northEastUp, northWestUp));
                        //lines.Add(Line.CreateBound(northEastUp, southEastUp));
                        //DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel)).SetShape(lines);

                    }
                }
            }

            //XYZ sbp1 = new XYZ(selectionBoundary.Left, selectionBoundary.Bottom, selectionBoundary.Down);
            //XYZ sbp2 = new XYZ(selectionBoundary.Left, selectionBoundary.Top, selectionBoundary.Down);
            //XYZ sbp3 = new XYZ(selectionBoundary.Right, selectionBoundary.Top, selectionBoundary.Down);
            //XYZ sbp4 = new XYZ(selectionBoundary.Right, selectionBoundary.Bottom, selectionBoundary.Down);

            //XYZ sbp5 = new XYZ(selectionBoundary.Left, selectionBoundary.Bottom, selectionBoundary.Up);
            //XYZ sbp6 = new XYZ(selectionBoundary.Left, selectionBoundary.Top, selectionBoundary.Up);
            //XYZ sbp7 = new XYZ(selectionBoundary.Right, selectionBoundary.Top, selectionBoundary.Up);
            //XYZ sbp8 = new XYZ(selectionBoundary.Right, selectionBoundary.Bottom, selectionBoundary.Up);


            //lines.Add(Line.CreateBound(sbp1, sbp2));
            //lines.Add(Line.CreateBound(sbp2, sbp3));
            //lines.Add(Line.CreateBound(sbp3, sbp4));
            //lines.Add(Line.CreateBound(sbp4, sbp1));

            //lines.Add(Line.CreateBound(sbp5, sbp6));
            //lines.Add(Line.CreateBound(sbp6, sbp7));
            //lines.Add(Line.CreateBound(sbp7, sbp8));
            //lines.Add(Line.CreateBound(sbp5, sbp8));

            //lines.Add(Line.CreateBound(sbp1, sbp5));
            //lines.Add(Line.CreateBound(sbp2, sbp6));
            //lines.Add(Line.CreateBound(sbp3, sbp7));
            //lines.Add(Line.CreateBound(sbp4, sbp8));





            tr.Commit();
            tr.Dispose();
            td(sb.ToString());
            #region Pipes
            foreach (Element pipe in pipes)
            {
                bool isFireFighting = false;
                WorksetId wsid = pipe.WorksetId;
                string ws = worksetnames[worksetIDs.IndexOf(wsid)];
                int Rank = GetSystemRank(ws);
                double newdia;
                try
                {

                    newdia = pipe.LookupParameter("Diameter").AsDouble() * 304.8; //mm
                }
                catch
                {
                    continue;
                }
                double spacing = SysSpacing(Rank, newdia);
                if (spacing == 0)
                {
                    continue;
                }
                Curve pipeCurve = ((LocationCurve)pipe.Location).Curve;
                double pipeOffset = 500 / 304.80;
                XYZ pipeDirection = ((Line)pipeCurve).Direction.Normalize();
                XYZ pipePerpendicular = new XYZ(-pipeDirection.Y, pipeDirection.X, pipeDirection.Z);
                XYZ P0 = pipeCurve.Evaluate(0, true);
                XYZ Pf = pipeCurve.Evaluate(1, true);
                XYZ Ps = P0.Add(ductOffset * pipeDirection);
                XYZ Pe = Pf.Add(-ductOffset * pipeDirection);
                Curve hangCurve = null;
                try
                {
                    hangCurve = Line.CreateBound(Ps, Pe) as Curve;
                }
                catch { continue; }

                if (Rank == 4)
                {
                    isFireFighting = true;
                }
                List<XYZ> pps = new List<XYZ>();
                List<XYZ> pipeHangPts = new List<XYZ>();
                double pipeDiam = (newdia / 304.8) + (2 * pipe.LookupParameter("Insulation Thickness").AsDouble());
                double hangerDiameter = -1;
                foreach (double hangerDiam in HangDias)
                {
                    if (pipeDiam <= (hangerDiam / 304.8))
                    {
                        hangerDiameter = hangerDiam / 304.8;
                        break;
                    }
                }
                if (hangerDiameter == -1) { hangerDiameter = 402 / 304.8; }
                ElementId levelId = pipe.LookupParameter("Reference Level").AsElementId();
                double outsideDiam = pipe.LookupParameter("Outside Diameter").AsDouble();
                double pipeElevation = pipeCurve.GetEndPoint(0).Z - pipe.LookupParameter("Insulation Thickness").AsDouble();
                double midElevStart = pipe.LookupParameter("Start Middle Elevation").AsDouble();
                double midElevEnd = pipe.LookupParameter("End Middle Elevation").AsDouble();
                double slope = pipe.LookupParameter("Slope").AsDouble();
                if (pipeCurve.Length > negLength && pipeCurve.Length <= pipeOffset)
                {
                    XYZ midPt = pipeCurve.Evaluate(0.50, true);
                    if (!pipeHangPts.Contains(midPt)) pipeHangPts.Add(midPt);
                }
                else if (pipeCurve.Length <= spacing && pipeCurve.Length > pipeOffset)
                {
                    if (!pipeHangPts.Contains(Ps)) pipeHangPts.Add(Ps);
                    if (!pipeHangPts.Contains(Pe)) pipeHangPts.Add(Pe);
                }
                else if (pipeCurve.Length > spacing)
                {
                    if (!pipeHangPts.Contains(Ps)) pipeHangPts.Add(Ps);
                    double n = Math.Ceiling(hangCurve.Length / spacing) - 1;
                    XYZ prev = Ps;
                    double Ns = (hangCurve.Length / (n + 1));
                    for (int i = 0; i < n; i++)
                    {
                        XYZ point = prev.Add(Ns * pipeDirection);
                        if (!pipeHangPts.Contains(point)) pipeHangPts.Add(point);
                        prev = point;
                    }
                    if (!pipeHangPts.Contains(Pe)) pipeHangPts.Add(Pe);
                }

                pipeHangers.Add(new PipeHanger(pipe, P0, isFireFighting, hangerDiameter, levelId, pipeElevation, midElevStart, midElevEnd, slope, pipeHangPts, pipeDirection, pipePerpendicular));
            }
            #endregion

            #region Cable Trays
            foreach (Element tray in cables)
            {
                Curve trayCurve = ((LocationCurve)tray.Location).Curve;
                double trayOffset = 500 / 304.80;
                XYZ trayDir = ((Line)trayCurve).Direction.Normalize();
                XYZ perpendicular = new XYZ(-trayDir.Y, trayDir.X, trayDir.Z);
                XYZ P0 = trayCurve.Evaluate(0, true);
                XYZ Pf = trayCurve.Evaluate(1, true);
                XYZ Ps = P0.Add(ductOffset * trayDir);
                XYZ Pe = Pf.Add(-ductOffset * trayDir);
                double spacing0 = 1500 / 304.8;
                Curve hangCurve = null;
                try
                {
                    hangCurve = Line.CreateBound(Ps, Pe) as Curve;
                }
                catch { continue; }
                List<XYZ> trayHangPts = new List<XYZ>();
                double width = tray.LookupParameter("Width").AsDouble();
                ElementId levelId = tray.LookupParameter("Reference Level").AsElementId();
                double botElevation = tray.LookupParameter("Bottom Elevation").AsDouble();
                double elevation = ((Level)doc.GetElement(levelId)).Elevation + botElevation;
                if (trayCurve.Length > negLength && trayCurve.Length <= trayOffset)
                {
                    XYZ midPt = trayCurve.Evaluate(0.50, true);
                    if (!trayHangPts.Contains(midPt)) trayHangPts.Add(midPt);
                }
                else if (trayCurve.Length <= spacing0 && trayCurve.Length > trayOffset)
                {
                    if (!trayHangPts.Contains(Ps)) trayHangPts.Add(Ps);
                    if (!trayHangPts.Contains(Pe)) trayHangPts.Add(Pe);
                }
                else if (trayCurve.Length > spacing0)
                {
                    if (!trayHangPts.Contains(Ps)) trayHangPts.Add(Ps);
                    double n = Math.Floor((hangCurve.Length + (100 / 304.8)) / spacing0);
                    XYZ prevPt = Ps;
                    for (int i = 0; i < n; i++)
                    {
                        XYZ point = prevPt.Add(spacing0 * trayDir);
                        if (!trayHangPts.Contains(point)) trayHangPts.Add(point);
                        prevPt = point;
                    }
                }
                trayHangers.Add(new TrayHanger(tray, width, levelId, elevation, trayHangPts, trayDir, perpendicular));
            }
            #endregion

            //====================================================================================================
            //====================================================================================================
            commitTransaction();
            return Result.Succeeded;
        }
        void commitTransaction()
        {
            string err = ""; int errco = 0;
            using (Transaction trans = new Transaction(doc, "IBIMS_Hangers"))
            {
                trans.Start();
                int h = 0;
                pipeHanger20.Activate();
                ductHanger.Activate();
                pipeHanger2.Activate();
                pipeHanger200.Activate();
                td("Transaction" + ductHangers.Count.ToString());
                foreach (DuctHanger hanger in ductHangers)
                {

                    if (hanger.supports.Count > 0)
                    {
                        foreach (Support support in hanger.supports)
                        {
                            FamilyInstance hang = doc.Create.NewFamilyInstance(support.point, ductHanger, hanger.ductPerpendicular, hanger.duct, StructuralType.NonStructural);
                            hang.LookupParameter("Width").Set(hanger.ductWidth + (2 * hanger.insoThick) + 16 / 304.8);
                            double Z = hanger.botElevation - hanger.insoThick - hang.LookupParameter("Elevation from Level").AsDouble();
                            hang.Location.Move(new XYZ(0, 0, Z));
                            support.rod += hanger.insoThick + (hanger.ductHeight / 2);
                            hang.LookupParameter("ROD 1").Set(support.rod);
                            hang.LookupParameter("ROD 2").Set(support.rod);
                        }
                    }
                    h++;
                }
                foreach (PipeHanger pipeHanger in pipeHangers)
                {
                    int hangco = 0;
                    foreach (XYZ point in pipeHanger.pipeHangPts)
                    {
                        int fflu = -1; int ffld = -1; double ddu = double.MinValue; double ddd = double.MaxValue;
                        double Zd = 0; double Zu = 0; XYZ IP = null;
                        foreach (Face face in floorFacesDown)
                        {
                            Line ll = Line.CreateUnbound(point, XYZ.BasisZ);
                            IntersectionResultArray iraa = new IntersectionResultArray();
                            SetComparisonResult scr = face.Intersect(ll, out iraa);
                            if (iraa != null)
                            {
                                if (!iraa.IsEmpty)
                                {
                                    XYZ ip = iraa.get_Item(0).XYZPoint;
                                    double D = point.Z - ip.Z;
                                    if (D > 0 && D <= ddd)
                                    {
                                        ddd = D;
                                        ffld = floorFacesDown.IndexOf(face); Zd = ip.Z; IP = ip;
                                    }
                                    else if (D < 0 && D >= ddu)
                                    {
                                        ddu = D;
                                        fflu = floorFacesDown.IndexOf(face); Zu = ip.Z; IP = ip;
                                    }
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            else
                            {
                                continue;
                            }
                        }
                        double hu = Math.Abs(Zu - point.Z);
                        double hd = Math.Abs(point.Z - Zd);
                        Face facee = null; int ffl = 0;
                        if (hu < hd && fflu != -1)
                        {
                            if (pipeHanger.isFireFighting)
                            {
                                FamilyInstance pangq = doc.Create.NewFamilyInstance(point, pipeHanger2, pipeHanger.pipePerpendicular, doc.GetElement(pipeHanger.levelId), StructuralType.NonStructural);
                                double q = 1;
                                if (pipeHanger.midElevEnd < pipeHanger.midElevStart)
                                {
                                    q = -1;
                                }
                                double ppeleve = pipeHanger.pipeElevation + (q * pipeHanger.slope * point.DistanceTo(pipeHanger.startPt));
                                double pofl = ppeleve - (3000 / 304.8);
                                pangq.LookupParameter("Diameter").Set(pipeHanger.hangerDiameter);
                                pangq.LookupParameter("Offset from Host").Set(pofl);
                                double pangelev = ppeleve - (3000 / 304.8);
                                if (floorElevations[fflu] > pangelev)
                                {
                                    //double ae = floorelevs[fflu] - pangelev;
                                    double ae = hu + (2995 / 304.8);
                                    pangq.LookupParameter("AnchorElevation").Set(ae);
                                }
                            }
                            else
                            {
                                facee = floorFacesDown[fflu]; ffl = fflu;
                            }
                        }
                        else if (hu > hd && ffld != -1)
                        {
                            facee = floorFacesUp[ffld]; ffl = ffld;
                        }
                        if (!pipeHanger.isFireFighting && facee != null)
                        {
                            Reference reffface = null;
                            if (Math.Round(Math.Abs(((PlanarFace)facee).FaceNormal.Z), 3) == 1)
                            {
                                if (facee.Reference.CreateLinkReference(RLI) != null) { reffface = facee.Reference.CreateLinkReference(RLI); }
                            }
                            else // Ramp
                            {
                                FilteredElementCollector airs = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol));
                                FamilySymbol fsair = null;
                                foreach (FamilySymbol fsa in airs)
                                {
                                    if (fsa.Name == "Air") { fsair = fsa; fsair.Activate(); }
                                }
                                FamilyInstance air = null;
                                air = doc.Create.NewFamilyInstance(IP, fsair, StructuralType.NonStructural);
                                doc.Regenerate();
                                Face faceair = GetFaces(air);
                                string refnew = air.UniqueId + ":0:INSTANCE:" + faceair.Reference.ConvertToStableRepresentation(doc);
                                reffface = Reference.ParseFromStableRepresentation(doc, refnew);
                            }
                            try
                            {
                                FamilySymbol FS = null;
                                if (pipeHanger.hangerDiameter > (202 / 304.8))
                                {
                                    FS = pipeHanger200;
                                }
                                else { FS = pipeHanger20; }
                                FS.Activate();
                                FamilyInstance pang = doc.Create.NewFamilyInstance(reffface, point, pipeHanger.pipeDirection, FS);
                                pang.LookupParameter("Schedule Level").Set(floors[ffl].LevelId);
                                Line ll = Line.CreateUnbound(point, XYZ.BasisZ);
                                double rr = pang.HandOrientation.AngleOnPlaneTo(pipeHanger.pipeDirection, XYZ.BasisZ);
                                IntersectionResultArray iraa = new IntersectionResultArray();
                                SetComparisonResult scr = facee.Intersect(ll, out iraa);
                                if (iraa != null && !iraa.IsEmpty)
                                {
                                    Curve cv = Line.CreateBound(point, iraa.get_Item(0).XYZPoint);
                                    pang.LookupParameter("Pipe_distance").Set(cv.Length - (0.5 * pipeHanger.hangerDiameter));
                                }
                                pang.LookupParameter("Pipe Outer Diameter").Set(pipeHanger.hangerDiameter);
                            }
                            catch
                            {
                                if (hangco == 0)
                                {
                                    errco++;
                                    err += pipeHanger.pipe.Id + "\n";
                                }
                            }
                        }
                        hangco++;
                    }
                }
                foreach (TrayHanger tray in trayHangers)
                {
                    foreach (XYZ p in tray.hangPts)
                    {
                        int fflu = -1;
                        double ddu = double.MinValue;
                        double Zu = 0;
                        foreach (Face face in floorFacesDown)
                        {
                            Line ll = Line.CreateUnbound(p, XYZ.BasisZ);
                            IntersectionResultArray iraa = new IntersectionResultArray();
                            SetComparisonResult scr = face.Intersect(ll, out iraa);
                            if (iraa != null)
                            {
                                if (!iraa.IsEmpty)
                                {
                                    XYZ ip = iraa.get_Item(0).XYZPoint;
                                    double D = p.Z - ip.Z;
                                    if (D < 0 && D >= ddu)
                                    {
                                        ddu = D;
                                        fflu = floorFacesDown.IndexOf(face); Zu = ip.Z;
                                    }
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            else
                            {
                                continue;
                            }
                        }
                        double hu = Math.Abs(Zu - p.Z); double ROD = hu;
                        if (fflu != -1)
                        {
                            XYZ PP = new XYZ(p.X, p.Y, tray.trayElevation); ductHanger.Activate();
                            FamilyInstance hang = doc.Create.NewFamilyInstance(PP, ductHanger, tray.trayPerpendicular, doc.GetElement(tray.levelId), StructuralType.NonStructural);
                            hang.LookupParameter("Width").Set(tray.trayWidth + 100 / 304.8);
                            hang.LookupParameter("ROD 1").Set(ROD);
                            hang.LookupParameter("ROD 2").Set(ROD);
                        }
                    }
                }
                if (errco > 0)
                {
                    TaskDialog.Show("Warning", errco + " Pipes have no Hangers." + Environment.NewLine + err);
                }
                trans.Commit();
                trans.Dispose();
            }
        }

        Face GetFaces(Element ele)
        {
            Face face = null;
            Solid s1 = null;
            options.ComputeReferences = true;
            GeometryElement gele = ele.get_Geometry(options);
            foreach (GeometryObject geo in gele)
            {
                GeometryInstance Gi = geo as GeometryInstance;
                foreach (GeometryObject gi in Gi.GetInstanceGeometry())
                {
                    Solid g = gi as Solid;
                    if (g != null && g.Volume != 0)
                    {
                        s1 = g;
                        break;
                    }
                }
            }
            if (s1 != null)
            {
                foreach (Face fa in s1.Faces)
                {
                    if (Math.Round(((PlanarFace)fa).FaceNormal.Z, 2) == -1)
                    {
                        face = fa; break;
                    }
                }
            }
            return face;
        }

        double GetSpace(List<double> dias, List<double> spacs, double DIA)
        {
            int co = 0; double spc = 0;
            foreach (double d in dias)
            {
                if (dias[0] == -10)
                {
                    spc = spacs[0];
                    break;
                }
                else if (Math.Round(DIA, 5) <= Math.Round(d, 5))
                {
                    if (Math.Round(DIA, 5) == Math.Round(d, 5))
                    {
                        spc = spacs[co];
                    }
                    else
                    {
                        if (co == 0)
                        {
                            spc = spacs[co];
                        }
                        else
                        {
                            spc = spacs[co - 1];
                        }
                    }
                    break;
                }
                co++;
            }
            if (spc == 0)
            {
                spc = spacs.Last();
            }
            return spc;
        }

        double SysSpacing(int rankk, double dia)
        {
            double spac = 0;
            if (rankk == 1)
            {
                spac = GetSpace(waterSupplyDiams, waterSupplySpaces, dia);
            }
            else if (rankk == 2)
            {
                spac = GetSpace(chilledWaterDiams, chilledWaterSpaces, dia);
            }
            else if (rankk == 3)
            {
                spac = GetSpace(drainageDiams, drainageSpaces, dia);
            }
            else if (rankk == 4)
            {
                spac = GetSpace(fireDiams, fireSpaces, dia);
            }
            return spac;
        }

        int GetSystemRank(string wst)
        {
            int R = -1; int a = 0;
            foreach (List<string> ls in AllWorksetNames)
            {
                foreach (string w in ls)
                {
                    if (w == wst)
                    {
                        R = a;
                        break;
                    }
                }
                a++;
            }
            return R;
        }

        List<XYZ> decOrder(List<XYZ> oldlist, Curve cu)
        {
            if (Math.Round(((Line)cu).Direction.Normalize().Y, 3) == 0)
            {
                return oldlist.OrderByDescending(a => a.X).ToList();
            }
            else
            {
                return oldlist.OrderByDescending(a => a.Y).ToList();
            }
        }
        private XYZ getLocation(Element elem)
        {
            Location location = elem.Location;
            if (location is LocationPoint)
            {
                return ((LocationPoint)location).Point;
            }
            else
            {
                Curve curve = ((LocationCurve)location).Curve;
                return curve.Evaluate(0.5, true);
            }
        }
        void td(string message)
        {
            TaskDialog.Show("Info", message);
        }

        public Solid getSolid(Element elem)
        {
            IList<Solid> solids = new List<Solid>();
            try
            {

                GeometryElement geo = elem.get_Geometry(options);
                if (geo.FirstOrDefault() is Solid)
                {
                    Solid solid = (Solid)geo.FirstOrDefault();
                    return SolidUtils.Clone(solid);
                }
                foreach (GeometryObject geometryObject in geo)
                {
                    if (geometryObject != null)
                    {
                        Solid solid = geometryObject as Solid;
                        if (solid != null && solid.Volume > 0)
                        {
                            solids.Add(solid);

                        }
                    }
                }
            }
            catch
            {
            }
            if (solids.Count == 0)
            {
                try
                {
                    GeometryElement geo = elem.get_Geometry(options);
                    GeometryInstance geoIns = geo.FirstOrDefault() as GeometryInstance;
                    if (geoIns != null)
                    {
                        GeometryElement geoElem = geoIns.GetInstanceGeometry();
                        if (geoElem != null)
                        {
                            foreach (GeometryObject geometryObject in geoElem)
                            {
                                Solid solid = geometryObject as Solid;
                                if (solid != null && solid.Volume > 0)
                                {
                                    solids.Add(solid);
                                }
                            }
                        }
                    }
                }
                catch
                {
                    throw new InvalidOperationException();
                }
            }
            if (solids.Count > 0)
            {
                try
                {

                    return SolidUtils.Clone(solids.OrderByDescending(x => x.Volume).ElementAt(0));
                }
                catch
                {
                    return solids.OrderByDescending(x => x.Volume).ElementAt(0);
                }
            }
            else
            {
                return null;
            }
        }

        public void processDucts(List<Element> ducts, List<Element> ductFits, Face floorFaceDown)
        {

            #region Ducts
            foreach (Element duct in ducts)
            {
                double spacingFin = 0;
                Curve ductCurve = ((LocationCurve)duct.Location).Curve;
                XYZ ductOrigin = ductCurve.Evaluate(0.5, true);
                double ductWidth = 0;
                List<XYZ> pts = new List<XYZ>();
                try { ductWidth = duct.LookupParameter("Width").AsDouble(); }
                catch { ductWidth = duct.LookupParameter("Diameter").AsDouble(); }
                if (AllWorksetsDIMS[0].Count == 2)
                {
                    spacingFin = Convert.ToDouble(AllWorksetsDIMS[0][1][0]);
                }
                else
                {
                    for (int j = 0; j < AllWorksetsDIMS[0][0].Count; j++)
                    {
                        double from = Convert.ToDouble(AllWorksetsDIMS[0][0][j]);
                        double to = Convert.ToDouble(AllWorksetsDIMS[0][1][j]);
                        double spacing = Convert.ToDouble(AllWorksetsDIMS[0][2][j]);
                        if (ductWidth * 304.8 > from && ductWidth * 304.8 <= to)
                        {
                            spacingFin = spacing / 304.8;
                            break;
                        }
                    }
                }
                if (spacingFin == 0)
                {
                    continue;
                }
                ElementId levelId = duct.LookupParameter("Reference Level").AsElementId();
                double ductHeight = 0;
                try { ductHeight = duct.LookupParameter("Height").AsDouble(); }
                catch { ductHeight = duct.LookupParameter("Diameter").AsDouble(); }
                double botElevation = duct.LookupParameter("Bottom Elevation").AsDouble() + ((Level)doc.GetElement(levelId)).Elevation;
                double insoThick = duct.LookupParameter("Insulation Thickness").AsDouble();
                XYZ ductDir = ((Line)ductCurve).Direction.Normalize();
                XYZ ductPerpendicular = new XYZ(-ductDir.Y, ductDir.X, ductDir.Z);
                XYZ ductMidPt = ductCurve.Evaluate(0.5, true);

                #region Duct Fitting
                // getting the duct fitting points that intersects the curve line.
                // then offsets the hangers by a margin that insures the hangers will never intersect the fitting.

                foreach (Element fitting in ductfits)
                {
                    double angle = 0;
                    double takeOffLength = 0;
                    Parameter width1Param = fitting.LookupParameter("Duct Width 1");
                    Parameter takeOffParam = fitting.LookupParameter("Takeoff Fixed Length");
                    Parameter angleParam = fitting.LookupParameter("Angle");
                    FamilyInstance fittingFI = fitting as FamilyInstance;
                    if (width1Param == null)
                    {
                        continue;
                    }
                    try
                    {
                        takeOffLength = takeOffParam.AsDouble();
                        angle = angleParam.AsDouble();
                    }
                    catch { }
                    double radius = Math.Sqrt(Math.Pow(ductWidth, 2) + Math.Pow(width1Param.AsDouble(), 2)) * 0.50;
                    XYZ fittingLocPt = ((LocationPoint)fitting.Location).Point;
                    double margin = 0.50 * (Math.Tan(angle) * takeOffLength);
                    XYZ center = new XYZ(fittingLocPt.X, fittingLocPt.Y, ductMidPt.Z).Add(-margin * fittingFI.FacingOrientation);
                    Curve circ = Ellipse.CreateCurve(center, radius, radius, ductDir, ductPerpendicular, 0, 2 * Math.PI * radius);
                    circ.Intersect(ductCurve, out IntersectionResultArray ira);
                    if (ira != null)
                    {
                        if (ira.Size == 2)
                        {
                            pts.Add(ira.get_Item(0).XYZPoint);
                            pts.Add(ira.get_Item(1).XYZPoint);
                        }
                    }
                }
                #endregion

                XYZ P0 = ductCurve.Evaluate(0, true);
                XYZ Pf = ductCurve.Evaluate(1, true);
                List<XYZ> fittingPts = decOrder(pts, ductCurve);
                List<XYZ> ductCurvePts = decOrder(new List<XYZ>() { P0, Pf }, ductCurve);
                XYZ dir = Line.CreateBound(ductCurvePts[0], ductCurvePts[1]).Direction.Normalize();
                XYZ Ps = ductCurvePts[0].Add(ductOffset * dir);
                XYZ Pe = ductCurvePts[1].Add(-ductOffset * dir);
                Curve cc = null;
                try
                {
                    cc = Line.CreateBound(Ps, Pe) as Curve;
                }
                catch { continue; }
                List<Support> Supports = new List<Support>();
                if (fittingPts.Count == 0) // No ductfittings
                {
                    if (ductCurve.Length > negLength && ductCurve.Length <= 4 * ductOffset)
                    {
                        double rod = getRod(ductMidPt, floorFaceDown);
                        XYZ PP = new XYZ(ductMidPt.X, ductMidPt.Y, ductMidPt.Z - insoThick - (ductHeight / 2));

                        if (rod != 0) Supports.Add(new Support(PP, rod));
                    }
                    else if (ductCurve.Length <= spacingFin && ductCurve.Length > 4 * ductOffset)
                    {
                        double rod1 = getRod(Ps, floorFaceDown);
                        XYZ ps = new XYZ(Ps.X, Ps.Y, Ps.Z - insoThick - (ductHeight / 2));
                        if (rod1 != 0) Supports.Add(new Support(ps, rod1));

                        double rod2 = getRod(Pe, floorFaceDown);
                        XYZ pe = new XYZ(Pe.X, Pe.Y, Pe.Z - insoThick - (ductHeight / 2));
                        if (rod2 != 0) Supports.Add(new Support(pe, rod2));
                    }
                    else if (ductCurve.Length > spacingFin) // collect point of hangers
                    {
                        double rod = getRod(Ps, floorFaceDown);
                        XYZ ps = new XYZ(Ps.X, Ps.Y, Ps.Z - insoThick - (ductHeight / 2));
                        if (rod != 0) Supports.Add(new Support(ps, rod));
                        double n = Math.Ceiling(cc.Length / spacingFin) - 1;
                        XYZ prevPt = Ps;
                        for (int i = 0; i < n; i++)
                        {
                            XYZ point = prevPt.Add(spacingFin * dir);
                            double ROD = getRod(point, floorFaceDown);
                            XYZ p = new XYZ(point.X, point.Y, point.Z - insoThick - (ductHeight / 2));
                            if (ROD != 0) Supports.Add(new Support(p, ROD));
                            prevPt = point;
                        }
                        double rod2 = getRod(Pe, floorFaceDown);
                        XYZ pe = new XYZ(Pe.X, Pe.Y, Pe.Z - insoThick - (ductHeight / 2));
                        if (rod2 != 0) Supports.Add(new Support(pe, rod2));
                    }
                }
                else       // With ductfittings
                {
                    if (ductCurve.Length > negLength && ductCurve.Length <= 4 * ductOffset)
                    {
                        List<XYZ> fitPtsOrdered = decOrder(new List<XYZ>() { fittingPts[0], fittingPts[1], ductMidPt }, ductCurve);
                        if (fitPtsOrdered.IndexOf(ductMidPt) == 1)
                        {
                            double rod = getRod(fitPtsOrdered[2], floorFaceDown);
                            XYZ point = new XYZ(fitPtsOrdered[2].X, fitPtsOrdered[2].Y, fitPtsOrdered[2].Z - insoThick - (ductHeight / 2));
                            if (rod != 0) Supports.Add(new Support(point, rod)); // farther point
                        }
                        else
                        {
                            double rod = getRod(ductMidPt, floorFaceDown);
                            XYZ point = new XYZ(ductMidPt.X, ductMidPt.Y, ductMidPt.Z - insoThick - (ductHeight / 2));
                            if (rod != 0) Supports.Add(new Support(point, rod)); // duct center
                        }
                    }
                    else if (ductCurve.Length <= spacingFin && ductCurve.Length > 4 * ductOffset)
                    {
                        List<XYZ> fitPtsOrdered = decOrder(new List<XYZ>() { fittingPts[0], fittingPts[1], Ps }, ductCurve);
                        if (fitPtsOrdered.IndexOf(Ps) == 1)
                        {
                            double rod = getRod(fitPtsOrdered[2], floorFaceDown);
                            XYZ point = new XYZ(fitPtsOrdered[2].X, fitPtsOrdered[2].Y, fitPtsOrdered[2].Z - insoThick - (ductHeight / 2));
                            if (rod != 0) Supports.Add(new Support(point, rod));
                        }
                        else
                        {
                            double rod = getRod(Ps, floorFaceDown);
                            XYZ point = new XYZ(Ps.X, Ps.Y, Ps.Z - insoThick - (ductHeight / 2));
                            if (rod != 0) Supports.Add(new Support(point, rod));
                        }
                        List<XYZ> ptss = new List<XYZ>() { Ps, Pe };
                        for (int i = 0; i < 2; i++)
                        {
                            for (int j = 0; j < fittingPts.Count; j += 2)
                            {
                                List<XYZ> pso = decOrder(new List<XYZ>() { fittingPts[j], fittingPts[j + 1], ptss[i] }, ductCurve);
                                if (pso.IndexOf(ptss[i]) == 1) //Between Case
                                {
                                    if (i == 0)
                                    {
                                        double rod = getRod(pso[2], floorFaceDown);
                                        XYZ point = new XYZ(pso[2].X, pso[2].Y, pso[2].Z - insoThick - (ductHeight / 2));
                                        if (rod != 0) Supports.Add(new Support(point, rod));
                                    }
                                    else
                                    {
                                        double rod = getRod(pso[0], floorFaceDown);
                                        XYZ point = new XYZ(pso[0].X, pso[0].Y, pso[0].Z - insoThick - (ductHeight / 2));
                                        if (rod != 0) Supports.Add(new Support(point, rod));
                                    }
                                    break;
                                }
                            }
                        }
                        double rod2 = getRod(Pe, floorFaceDown);
                        XYZ pe = new XYZ(Pe.X, Pe.Y, Pe.Z - insoThick - (ductHeight / 2));
                        if (rod2 != 0) Supports.Add(new Support(pe, rod2));
                    }
                    else if (ductCurve.Length > spacingFin)
                    {
                        List<XYZ> ps1o = decOrder(new List<XYZ>() { fittingPts[0], fittingPts[1], Ps }, ductCurve);
                        if (ps1o.IndexOf(Ps) == 1)
                        {
                            double rod = getRod(ps1o[2], floorFaceDown);
                            XYZ point = new XYZ(ps1o[2].X, ps1o[2].Y, ps1o[2].Z - insoThick - (ductHeight / 2));
                            if (rod != 0) Supports.Add(new Support(point, rod));
                        }
                        else
                        {
                            double rod = getRod(Ps, floorFaceDown);
                            XYZ ps = new XYZ(Ps.X, Ps.Y, Ps.Z - insoThick - (ductHeight / 2));
                            if (rod != 0) Supports.Add(new Support(ps, rod));
                        }
                        double n = Math.Ceiling(cc.Length / spacingFin) - 1;
                        XYZ prevPt = Ps;
                        for (int i = 0; i < n; i++)
                        {
                            XYZ point = prevPt.Add(spacingFin * dir);
                            for (int j = 0; j < fittingPts.Count; j += 2)
                            {
                                List<XYZ> ps3o = decOrder(new List<XYZ>() { fittingPts[j], fittingPts[j + 1], point }, ductCurve);
                                if (ps3o.IndexOf(point) == 1) //Between Case
                                {
                                    point = ps3o[0];
                                    break;
                                }
                            }
                            double rod = getRod(point, floorFaceDown);
                            XYZ P = new XYZ(point.X, point.Y, point.Z - insoThick - (ductHeight / 2));
                            if (rod != 0) Supports.Add(new Support(point, rod));
                            prevPt = point;
                        }
                    }
                }

                ductHangers.Add(new DuctHanger(duct, ductPerpendicular, ductWidth, ductHeight, insoThick, botElevation, Supports));

            }
            #endregion
        }

        private double getRod(XYZ p, Face floorFaceDown)
        {
            int fflu = -1;
            int ffld = -1;
            double ROD = 0;
            double dd = double.MinValue;
            double dd2 = double.MaxValue;
            double Zu = 0; double Zd = 0;
            Face face = floorFaceDown;

            Line ll = Line.CreateUnbound(p, XYZ.BasisZ);
            IntersectionResultArray iraa = new IntersectionResultArray();
            SetComparisonResult scr = face.Intersect(ll, out iraa);
            if (iraa == null || iraa.IsEmpty) return 0;

            List<double> FXs = new List<double>(); List<double> FYs = new List<double>();
            XYZ ip = iraa.get_Item(0).XYZPoint;
            double D = p.Z - ip.Z;
            if (D < 0 && D >= dd)
            {
                dd = D;
                fflu = floorFacesDown.IndexOf(face); Zu = ip.Z - p.Z;
            }
            else if (D > 0 && D <= dd2)
            {
                if (D < dd2)
                {
                    dd2 = D;
                    ffld = floorFacesDown.IndexOf(face); Zd = p.Z - ip.Z;
                }
            }


            if (fflu != -1 && ffld != -1)
            {
                ROD = Math.Min(Zu, Zd);
            }
            else if (fflu != -1)
            {
                ROD = Zu;
            }
            else if (ffld != -1)
            {
                ROD = Zd;
            }
            else
            {
                ROD = 2000 / 304.8;
            }
            return ROD;
        }
    }
}
