using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
        List<List<Dictionary<string, double>>> AllWorksetsDIMS;
        IList<Element> MechanicalEquipment, ducts, pipes, cables, floors, ductfits;
        IList<Reference> mechRefs, linkedRefs;
        List<double> HangDias;
        List<Workset> worksets;
        List<WorksetId> worksetIDs;
        Options options;
        RevitLinkInstance RLI;
        List<DuctHanger> ductHangers;
        List<PipeHanger> pipeHangers;
        List<TrayHanger> trayHangers;
        double meanFloorHeight, ductOffset, negLength;
        StringBuilder sb;
        double minx = double.MaxValue, miny = double.MaxValue, minz = double.MaxValue;
        double maxx = double.MinValue, maxy = double.MinValue, maxz = double.MinValue;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet element)
        {
            sb = new StringBuilder();
            Stopwatch watch = new Stopwatch();
            watch.Start();
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            options = new Options();
            options.ComputeReferences = true;
            options.View = doc.ActiveView;

            ductHangers = new List<DuctHanger>();
            pipeHangers = new List<PipeHanger>();
            trayHangers = new List<TrayHanger>();
            mechRefs = new List<Reference>();
            linkedRefs = new List<Reference>();
            ducts = new List<Element>();
            pipes = new List<Element>();
            cables = new List<Element>();
            floors = new List<Element>();

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
                return Result.Failed;
            }
            watch.Stop();
            sb.AppendLine($"check time for ramilies is {watch.ElapsedMilliseconds} ");
            watch.Restart();
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
            //td((meanFloorHeight * 304.8 / 1000).ToString());
            worksets = new FilteredWorksetCollector(doc).Where(x => x.Kind == WorksetKind.UserWorkset).ToList();
            worksetnames = worksets.Select(x => x.Name).ToList();
            worksetIDs = worksets.Select(x => x.Id).ToList();

            if (worksetnames.Count == 0)
            {
                TaskDialog.Show("Error", "Document has no UserWorksets.");
                return Result.Failed;
            }
            UI = new HangersFM(LinksNames, levelsNames);
            UI.ShowDialog();
            if (UI.canc)
            {
                return Result.Cancelled;
            }



            AllWorksetNames = UI.AllworksetsNames;
            AllWorksetsDIMS = UI.AllworksetsDIMS;



            RLI = linksFEC.Cast<RevitLinkInstance>()
                .Where(x => ((RevitLinkType)doc.GetElement(x.GetTypeId())).Name == LinksNames[UI.linkIndex]).First();
            LinkDoc = RLI.GetLinkDocument();


            watch.Stop();
            sb.AppendLine($"time for calculating info {watch.ElapsedMilliseconds}");
            watch.Restart();
            try
            {
                mechRefs = uidoc.Selection.PickObjects(ObjectType.Element, new SelectionFilterPDC(), "Select Pipes / Ducts / CableTrays.");
                foreach (Reference reference in mechRefs)
                {
                    Element e = doc.GetElement(reference);
                    if (e.Category.Name == "Ducts")
                    {
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
            watch.Stop();
            sb.AppendLine($"time to get selection{watch.ElapsedMilliseconds}");
            watch.Restart();
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

            //List<ElementId> ids = new List<ElementId>();
            ductfits = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DuctFitting)
                .OfClass(typeof(FamilyInstance))
                //.Where(x => selectionBoundary.contains(x))
                .ToList();

            getQuadTreeDims();
            QuadTree ductTree = new QuadTree(minx, maxy, maxx, miny, maxz, minz);
            QuadTree fitsTree = new QuadTree(minx, maxy, maxx, miny, maxz, minz);
            QuadTree pipesTree = new QuadTree(minx, maxy, maxx, miny, maxz, minz);
            QuadTree traysTree = new QuadTree(minx, maxy, maxx, miny, maxz, minz);

            foreach (Element duct in ducts)
            {
                ductTree.insert(duct);
            }
            foreach (Element fitting in ductfits)
            {
                fitsTree.insert(fitting);
            }
            foreach (Element pipe in pipes)
            {
                pipesTree.insert(pipe);
            }
            td($"Cable trays count is :{cables.Count}");
            double count = 0;
            foreach (Element tray in cables)
            {
                if (traysTree.insert(tray)) count++;
            }
            td($"Count: {count}");
            watch.Stop();
            sb.AppendLine($"time to quadtree elements {watch.ElapsedMilliseconds}");
            //Transaction tr = new Transaction(doc, "draw");
            //tr.Start();
            foreach (Element ele in floors)
            {
                watch.Restart();
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
                                    if (Math.Max(p1.Y, p2.Y) > floorTop) floorTop = Math.Max(p1.Y, p2.Y);
                                    if (Math.Max(p1.X, p2.X) > floorRight) floorRight = Math.Max(p1.X, p2.X);
                                    if (Math.Min(p1.Y, p2.Y) < floorBottom) floorBottom = Math.Min(p1.Y, p2.Y);
                                    if (Math.Max(p1.Z, p2.Z) > floorUp) floorUp = Math.Max(p1.Z, p2.Z);
                                    if (Math.Min(p1.Z, p2.Z) < floorDown) floorDown = Math.Min(p1.Z, p2.Z);
                                }

                            }
                        }
                        floorRight += Math.Max((meanFloorHeight / 2), 2000 / 304.8);
                        floorLeft -= Math.Max((meanFloorHeight / 2), 2000 / 304.8);
                        floorTop += Math.Max((meanFloorHeight / 2), 2000 / 304.8);
                        floorBottom -= Math.Max((meanFloorHeight / 2), 2000 / 304.8);
                        floorUp += Math.Max((meanFloorHeight / 2), 2000 / 304.8);
                        floorDown -= Math.Max((meanFloorHeight / 2), 2000 / 304.8);
                        //Boundary floorRange = new Boundary(double.MinValue, double.MaxValue, double.MaxValue, double.MinValue, double.MaxValue, double.MinValue);
                        Boundary floorRange = new Boundary(floorLeft, floorTop, floorRight, floorBottom, floorUp, floorDown);
                        List<Element> ductsInRange = ductTree.query(floorRange);
                        List<Element> fitsInRange = fitsTree.query(floorRange);
                        List<Element> pipesInRange = pipesTree.query(floorRange);
                        List<Element> traysInRange = traysTree.query(floorRange);

                        watch.Stop();
                        sb.AppendLine($"floor loop {watch.ElapsedMilliseconds}");
                        watch.Restart();
                        processDucts(ductsInRange, fitsInRange, solid.Faces, floorUp, floorDown);
                        watch.Stop();
                        sb.AppendLine($"Proccess ducts end {watch.ElapsedMilliseconds}");
                        watch.Restart();
                        processPipes(ele as Floor, pipesInRange, solid.Faces, floorUp, floorDown, elevationAtBot);
                        watch.Stop();
                        sb.AppendLine($"Proccess Pipes end {watch.ElapsedMilliseconds}");
                        watch.Restart();
                        processTrays(traysInRange, solid.Faces, floorUp, floorDown);
                        watch.Stop();
                        sb.AppendLine($"Proccess trays end {watch.ElapsedMilliseconds}");
                    }
                }
            }
            //doc.ActiveView.IsolateElementsTemporary(ids);
            //tr.Commit();
            //tr.Dispose();

            //====================================================================================================
            //====================================================================================================
            td(sb.ToString());
            commitTransaction();
            return Result.Succeeded;
        }

        private void getQuadTreeDims()
        {
            List<double> minxGen = new List<double>(), minyGen = new List<double>(), minzGen = new List<double>(), maxxGen = new List<double>(), maxyGen = new List<double>(), maxzGen = new List<double>();
            if (ductfits.Count > 0)
            {
                minxGen.Add(ductfits.Select(x => ((LocationPoint)x.Location).Point.X).Min());
                minyGen.Add(ductfits.Select(x => ((LocationPoint)x.Location).Point.Y).Min());
                minzGen.Add(ductfits.Select(x => ((LocationPoint)x.Location).Point.Z).Min());
                maxxGen.Add(ductfits.Select(x => ((LocationPoint)x.Location).Point.X).Max());
                maxyGen.Add(ductfits.Select(x => ((LocationPoint)x.Location).Point.Y).Max());
                maxzGen.Add(ductfits.Select(x => ((LocationPoint)x.Location).Point.Z).Max());
            }
            if (pipes.Count > 0)
            {
                List<XYZ> pipesP0 = pipes.Select(p => ((LocationCurve)p.Location).Curve.GetEndPoint(0)).ToList();
                List<XYZ> pipesP1 = pipes.Select(p => ((LocationCurve)p.Location).Curve.GetEndPoint(1)).ToList();
                minxGen.Add(Math.Min(pipesP0.Select(x => x.X).Min(), pipesP1.Select(x => x.X).Min()));
                minyGen.Add(Math.Min(pipesP0.Select(x => x.Y).Min(), pipesP1.Select(x => x.Y).Min()));
                minzGen.Add(Math.Min(pipesP0.Select(x => x.Z).Min(), pipesP1.Select(x => x.Z).Min()));
                maxxGen.Add(Math.Max(pipesP0.Select(x => x.X).Max(), pipesP1.Select(x => x.X).Max()));
                maxyGen.Add(Math.Max(pipesP0.Select(x => x.Y).Max(), pipesP1.Select(x => x.Y).Max()));
                maxzGen.Add(Math.Max(pipesP0.Select(x => x.Z).Max(), pipesP1.Select(x => x.Z).Max()));
            }
            if (cables.Count > 0)
            {
                List<XYZ> traysP0 = cables.Select(p => ((LocationCurve)p.Location).Curve.GetEndPoint(0)).ToList();
                List<XYZ> traysP1 = cables.Select(p => ((LocationCurve)p.Location).Curve.GetEndPoint(1)).ToList();
                minxGen.Add(Math.Min(traysP0.Select(x => x.X).Min(), traysP1.Select(x => x.X).Min()));
                minyGen.Add(Math.Min(traysP0.Select(x => x.Y).Min(), traysP1.Select(x => x.Y).Min()));
                minzGen.Add(Math.Min(traysP0.Select(x => x.Z).Min(), traysP1.Select(x => x.Z).Min()));
                maxxGen.Add(Math.Max(traysP0.Select(x => x.X).Max(), traysP1.Select(x => x.X).Max()));
                maxyGen.Add(Math.Max(traysP0.Select(x => x.Y).Max(), traysP1.Select(x => x.Y).Max()));
                maxzGen.Add(Math.Max(traysP0.Select(x => x.Z).Max(), traysP1.Select(x => x.Z).Max()));
            }
            minx = minxGen.Min() - (meanFloorHeight / 2);
            miny = minyGen.Min() - (meanFloorHeight / 2);
            minz = minzGen.Min() - (meanFloorHeight / 2);
            maxx = maxxGen.Max() + (meanFloorHeight / 2);
            maxy = maxyGen.Max() + (meanFloorHeight / 2);
            maxz = maxzGen.Max() + (meanFloorHeight / 2);
        }

        void commitTransaction()
        {
            using (Transaction tr = new Transaction(doc, "Hangers"))
            {
                tr.Start();
                string err = ""; int errco = 0;

                pipeHanger20.Activate();
                ductHanger.Activate();
                pipeHanger2.Activate();
                pipeHanger200.Activate();
                //td("Transaction" + ductHangers.Count.ToString());
                #region ducts
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
                }
                #endregion

                #region pipes
                foreach (PipeHanger pipeHanger in pipeHangers)
                {
                    foreach (Support support in pipeHanger.supports)
                    {
                        if (pipeHanger.isFireFighting)
                        {
                            FamilyInstance hangerFamInst = doc.Create.NewFamilyInstance(support.point, pipeHanger2, pipeHanger.pipePerpendicular, doc.GetElement(pipeHanger.levelId), StructuralType.NonStructural);
                            double q = 1;
                            if (pipeHanger.midElevEnd < pipeHanger.midElevStart)
                            {
                                q = -1;
                            }
                            double elev = pipeHanger.pipeElevation + (q * pipeHanger.slope * support.point.DistanceTo(pipeHanger.startPt));
                            double pipeOffsetFromHost = elev - (3000 / 304.8);
                            hangerFamInst.LookupParameter("Diameter").Set(pipeHanger.hangerDiameter);
                            hangerFamInst.LookupParameter("Offset from Host").Set(pipeOffsetFromHost);
                            hangerFamInst.LookupParameter("AnchorElevation").Set(support.rod);
                        }
                        else
                        {
                            Reference reference = null;
                            if (Math.Round(Math.Abs(((PlanarFace)pipeHanger.lowerFace).FaceNormal.Z), 3) == 1)
                            {
                                if (pipeHanger.lowerFace.Reference.CreateLinkReference(RLI) != null) { reference = pipeHanger.lowerFace.Reference.CreateLinkReference(RLI); }
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
                                air = doc.Create.NewFamilyInstance(support.point, fsair, StructuralType.NonStructural);
                                doc.Regenerate();
                                Face faceair = getSolid(air).Faces.get_Item(0);
                                string refnew = air.UniqueId + ":0:INSTANCE:" + faceair.Reference.ConvertToStableRepresentation(doc);
                                reference = Reference.ParseFromStableRepresentation(doc, refnew);
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
                                FamilyInstance pang = doc.Create.NewFamilyInstance(reference, support.point, pipeHanger.pipeDirection, FS);
                                pang.LookupParameter("Schedule Level").Set(pipeHanger.levelId);
                                Line ll = Line.CreateUnbound(support.point, XYZ.BasisZ);
                                double rr = pang.HandOrientation.AngleOnPlaneTo(pipeHanger.pipeDirection, XYZ.BasisZ);
                                IntersectionResultArray iraa = new IntersectionResultArray();
                                SetComparisonResult scr = pipeHanger.lowerFace.Intersect(ll, out iraa);
                                if (iraa != null && !iraa.IsEmpty)
                                {
                                    Curve cv = Line.CreateBound(support.point, iraa.get_Item(0).XYZPoint);
                                    pang.LookupParameter("Pipe_distance").Set(cv.Length - (0.5 * pipeHanger.hangerDiameter));
                                }
                                pang.LookupParameter("Pipe Outer Diameter").Set(pipeHanger.hangerDiameter);
                            }
                            catch
                            {
                                errco++;
                                err += pipeHanger.pipe.Id + "\n";
                            }
                        }
                    }
                }
                #endregion

                #region trays
                foreach (TrayHanger tray in trayHangers)
                {
                    foreach (Support support in tray.supports)
                    {
                        ductHanger.Activate();
                        FamilyInstance hang = doc.Create.NewFamilyInstance(support.point, ductHanger, tray.trayPerpendicular, doc.GetElement(tray.levelId), StructuralType.NonStructural);
                        hang.LookupParameter("Width").Set(tray.trayWidth + 100 / 304.8);
                        hang.LookupParameter("ROD 1").Set(support.rod);
                        hang.LookupParameter("ROD 2").Set(support.rod);
                    }
                }
                #endregion

                if (errco > 0)
                {
                    TaskDialog.Show("Warning", errco + " Pipes don't have Hangers." + Environment.NewLine + err);
                }
                tr.Commit();
                tr.Dispose();
            }
        }

        private void processPipes(Floor floor, List<Element> pipesInRange, FaceArray faces, double floorUp, double floorDown, double elevationAtBot)
        {
            #region Pipes

            foreach (Element pipe in pipesInRange)
            {
                List<Support> pipeSupports = new List<Support>();
                #region pipe
                bool isFireFighting = false;

                WorksetId wsid = pipe.WorksetId;
                string ws = worksetnames[worksetIDs.IndexOf(wsid)];
                double newdia, spacing;
                try
                {

                    newdia = pipe.LookupParameter("Diameter").AsDouble() * 304.8; //mm
                }
                catch
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
                int Rank;
                try
                {
                    hangCurve = Line.CreateBound(Ps, Pe) as Curve;
                }
                catch { continue; }
                if (pipe is RevitLinkInstance)
                {
                    RevitLinkType rlt = doc.GetElement(((RevitLinkInstance)pipe).GetTypeId()) as RevitLinkType;
                    Rank = GetSystemRank(rlt.Name);
                    spacing = SysSpacing(AllWorksetsDIMS[Rank], newdia);
                    if (Rank == 4)
                    {
                        isFireFighting = true;
                    }
                }
                else
                {
                    spacing = SysSpacing(AllWorksetsDIMS.Where(x => x.First()["spacing"] != 0).First().ToList(), newdia);
                }
                if (spacing == 0)
                {
                    continue;
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
                    if (!pipeHangPts.Contains(midPt))
                    {
                        pipeHangPts.Add(midPt);
                        double rod = getRod(midPt, faces);
                        if (rod != 0) pipeSupports.Add(new Support(midPt, rod + 2995 / 304.8));
                    }
                }
                else if (pipeCurve.Length <= spacing && pipeCurve.Length > pipeOffset)
                {
                    if (!pipeHangPts.Contains(Ps))
                    {
                        pipeHangPts.Add(Ps);
                        double rod = getRod(Ps, faces);
                        if (rod != 0) pipeSupports.Add(new Support(Ps, rod + 2995 / 304.8));

                    }
                    if (!pipeHangPts.Contains(Pe))
                    {
                        pipeHangPts.Add(Pe);
                        double rod = getRod(Pe, faces);
                        if (rod != 0) pipeSupports.Add(new Support(Pe, rod + 2995 / 304.8));

                    }
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
                        if (!pipeHangPts.Contains(point))
                        {
                            pipeHangPts.Add(point);
                            double rod = getRod(point, faces);
                            if (rod != 0) pipeSupports.Add(new Support(point, rod + 2995 / 304.8));
                        }
                        prev = point;
                    }
                    if (!pipeHangPts.Contains(Pe))
                    {
                        pipeHangPts.Add(Pe);
                        double rod = getRod(Pe, faces);
                        if (rod != 0) pipeSupports.Add(new Support(Pe, rod + 2995 / 304.8));

                    }
                }
                #endregion
                pipeHangers.Add(new PipeHanger(pipe, P0, isFireFighting, hangerDiameter, levelId, pipeElevation, midElevStart, midElevEnd, slope, pipeSupports, pipeDirection, pipePerpendicular, faces.get_Item(0)));
            }
            #endregion
        }
        public void processDucts(List<Element> ductsInRange, List<Element> fitsInRange, FaceArray faces, double up, double down)
        {

            #region Ducts
            foreach (Element duct in ductsInRange)
            {

                double spacingFin = 0;
                Curve ductCurve = ((LocationCurve)duct.Location).Curve;
                double minx = Math.Min(ductCurve.GetEndPoint(0).X, ductCurve.GetEndPoint(1).X) - 5;
                double miny = Math.Min(ductCurve.GetEndPoint(0).Y, ductCurve.GetEndPoint(1).Y) - 5;
                double maxx = Math.Max(ductCurve.GetEndPoint(0).X, ductCurve.GetEndPoint(1).X) + 5;
                double maxy = Math.Max(ductCurve.GetEndPoint(0).Y, ductCurve.GetEndPoint(1).Y) + 5;
                Boundary region = new Boundary(minx, maxy, maxx, miny, up, down);
                XYZ ductOrigin = ductCurve.Evaluate(0.5, true);
                double ductWidth = 0;
                List<XYZ> pts = new List<XYZ>();
                try { ductWidth = duct.LookupParameter("Width").AsDouble(); }
                catch { ductWidth = duct.LookupParameter("Diameter").AsDouble(); }
                if (AllWorksetsDIMS[0][0].Count == 1)
                {
                    spacingFin = AllWorksetsDIMS[0][0]["spacing"] / 304.8;
                }
                else
                {
                    if (AllWorksetsDIMS[0].Where(x => x["from"] < ductWidth * 304.8 && ductWidth * 304.8 <= x["to"]).Any())
                        spacingFin = AllWorksetsDIMS[0].Where(x => x["from"] < ductWidth * 304.8 && ductWidth * 304.8 <= x["to"]).First()["spacing"] / 304.8;
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

                List<Element> nearFits = fitsInRange.Where(x => region.contains(x)).ToList();
                foreach (Element fitting in nearFits)
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
                        double rod = getRod(ductMidPt, faces);
                        XYZ PP = new XYZ(ductMidPt.X, ductMidPt.Y, ductMidPt.Z - insoThick - (ductHeight / 2));

                        if (rod != 0) Supports.Add(new Support(PP, rod));
                    }
                    else if (ductCurve.Length <= spacingFin && ductCurve.Length > 4 * ductOffset)
                    {
                        double rod1 = getRod(Ps, faces);
                        XYZ ps = new XYZ(Ps.X, Ps.Y, Ps.Z - insoThick - (ductHeight / 2));
                        if (rod1 != 0) Supports.Add(new Support(ps, rod1));

                        double rod2 = getRod(Pe, faces);
                        XYZ pe = new XYZ(Pe.X, Pe.Y, Pe.Z - insoThick - (ductHeight / 2));
                        if (rod2 != 0) Supports.Add(new Support(pe, rod2));
                    }
                    else if (ductCurve.Length > spacingFin) // collect point of hangers
                    {
                        double rod = getRod(Ps, faces);
                        XYZ ps = new XYZ(Ps.X, Ps.Y, Ps.Z - insoThick - (ductHeight / 2));
                        if (rod != 0) Supports.Add(new Support(ps, rod));
                        double n = Math.Ceiling(cc.Length / spacingFin) - 1;
                        XYZ prevPt = Ps;
                        for (int i = 0; i < n; i++)
                        {
                            XYZ point = prevPt.Add(spacingFin * dir);
                            double ROD = getRod(point, faces);
                            XYZ p = new XYZ(point.X, point.Y, point.Z - insoThick - (ductHeight / 2));
                            if (ROD != 0) Supports.Add(new Support(p, ROD));
                            prevPt = point;
                        }
                        double rod2 = getRod(Pe, faces);
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
                            double rod = getRod(fitPtsOrdered[2], faces);
                            XYZ point = new XYZ(fitPtsOrdered[2].X, fitPtsOrdered[2].Y, fitPtsOrdered[2].Z - insoThick - (ductHeight / 2));
                            if (rod != 0) Supports.Add(new Support(point, rod)); // farther point
                        }
                        else
                        {
                            double rod = getRod(ductMidPt, faces);
                            XYZ point = new XYZ(ductMidPt.X, ductMidPt.Y, ductMidPt.Z - insoThick - (ductHeight / 2));
                            if (rod != 0) Supports.Add(new Support(point, rod)); // duct center
                        }
                    }

                    else if (ductCurve.Length <= spacingFin && ductCurve.Length > 4 * ductOffset)
                    {

                        List<XYZ> fitPtsOrdered = decOrder(new List<XYZ>() { fittingPts[0], fittingPts[1], Ps }, ductCurve);
                        if (fitPtsOrdered.IndexOf(Ps) == 1)
                        {
                            double rod = getRod(fitPtsOrdered[2], faces);
                            XYZ point = new XYZ(fitPtsOrdered[2].X, fitPtsOrdered[2].Y, fitPtsOrdered[2].Z - insoThick - (ductHeight / 2));
                            if (rod != 0) Supports.Add(new Support(point, rod));
                        }
                        else
                        {
                            double rod = getRod(Ps, faces);
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
                                        double rod = getRod(pso[2], faces);
                                        XYZ point = new XYZ(pso[2].X, pso[2].Y, pso[2].Z - insoThick - (ductHeight / 2));
                                        if (rod != 0) Supports.Add(new Support(point, rod));
                                    }
                                    else
                                    {
                                        double rod = getRod(pso[0], faces);
                                        XYZ point = new XYZ(pso[0].X, pso[0].Y, pso[0].Z - insoThick - (ductHeight / 2));
                                        if (rod != 0) Supports.Add(new Support(point, rod));
                                    }
                                    break;
                                }
                            }
                        }
                        double rod2 = getRod(Pe, faces);
                        XYZ pe = new XYZ(Pe.X, Pe.Y, Pe.Z - insoThick - (ductHeight / 2));
                        if (rod2 != 0) Supports.Add(new Support(pe, rod2));
                    }
                    else if (ductCurve.Length > spacingFin)
                    {
                        List<XYZ> ps1o = decOrder(new List<XYZ>() { fittingPts[0], fittingPts[1], Ps }, ductCurve);
                        if (ps1o.IndexOf(Ps) == 1)
                        {
                            double rod = getRod(ps1o[2], faces);
                            XYZ point = new XYZ(ps1o[2].X, ps1o[2].Y, ps1o[2].Z - insoThick - (ductHeight / 2));
                            if (rod != 0) Supports.Add(new Support(point, rod));
                        }
                        else
                        {
                            double rod = getRod(Ps, faces);
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
                            double rod = getRod(point, faces);
                            XYZ P = new XYZ(point.X, point.Y, point.Z - insoThick - (ductHeight / 2));
                            if (rod != 0) Supports.Add(new Support(point, rod));
                            prevPt = point;
                        }
                    }
                }
                List<Support> distinctSups = new List<Support>();
                foreach (Support sup in Supports)
                {
                    Support found = null;
                    found = distinctSups.Where(x => x.point.IsAlmostEqualTo(sup.point))?.FirstOrDefault();
                    if (found != null)
                    {
                        if (found.rod > sup.rod) found.rod = sup.rod;
                    }
                    else
                    {
                        distinctSups.Add(sup);
                    }
                }
                ductHangers.Add(new DuctHanger(duct, ductPerpendicular, ductWidth, ductHeight, insoThick, botElevation, distinctSups));

            }
            #endregion
        }

        public void processTrays(List<Element> traysInRange, FaceArray faces, double floorUp, double floorDown)
        {
            #region Cable Trays
            foreach (Element tray in traysInRange)
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
                List<Support> supports = new List<Support>();
                double width = tray.LookupParameter("Width").AsDouble();
                ElementId levelId = tray.LookupParameter("Reference Level").AsElementId();
                double botElevation = tray.LookupParameter("Bottom Elevation").AsDouble();
                double elevation = ((Level)doc.GetElement(levelId)).Elevation + botElevation;
                if (trayCurve.Length > negLength && trayCurve.Length <= trayOffset)
                {
                    XYZ midPt = trayCurve.Evaluate(0.50, true);
                    if (!trayHangPts.Contains(midPt))
                    {
                        trayHangPts.Add(midPt);
                        double rod = getRod(midPt, faces);
                        XYZ PP = new XYZ(midPt.X, midPt.Y, elevation);
                        if (rod != 0) supports.Add(new Support(PP, rod));
                    }
                }
                else if (trayCurve.Length <= spacing0 && trayCurve.Length > trayOffset)
                {
                    if (!trayHangPts.Contains(Ps))
                    {
                        trayHangPts.Add(Ps);
                        double rod = getRod(Ps, faces);
                        XYZ PP = new XYZ(Ps.X, Ps.Y, elevation);
                        if (rod != 0) supports.Add(new Support(PP, rod));
                    }
                    if (!trayHangPts.Contains(Pe))
                    {
                        trayHangPts.Add(Pe);
                        double rod = getRod(Pe, faces);
                        XYZ PP = new XYZ(Pe.X, Pe.Y, elevation);
                        if (rod != 0) supports.Add(new Support(PP, rod));
                    }
                }
                else if (trayCurve.Length > spacing0)
                {
                    if (!trayHangPts.Contains(Ps))
                    {
                        trayHangPts.Add(Ps);
                        double rod = getRod(Ps, faces);
                        XYZ PP = new XYZ(Ps.X, Ps.Y, elevation);
                        if (rod != 0) supports.Add(new Support(PP, rod));
                    }
                    double n = Math.Floor((hangCurve.Length + (100 / 304.8)) / spacing0);
                    XYZ prevPt = Ps;
                    for (int i = 0; i < n; i++)
                    {
                        XYZ point = prevPt.Add(spacing0 * trayDir);
                        if (!trayHangPts.Contains(point))
                        {
                            trayHangPts.Add(point);
                            double rod = getRod(point, faces);
                            XYZ PP = new XYZ(point.X, point.Y, elevation);
                            if (rod != 0) supports.Add(new Support(PP, rod));
                        }
                        prevPt = point;
                    }
                }

                trayHangers.Add(new TrayHanger(tray, width, levelId, supports, trayDir, perpendicular));
            }
            #endregion
        }
        double SysSpacing(List<Dictionary<string, double>> dims, double diameter)
        {
            if (dims == null) return 0;
            if (dims.Count == 0) return 0;
            else if (dims.Count == 1) return dims[0]["spacing"];
            else if (dims.Where(x => Math.Round(diameter, 5) <= Math.Round(x["size"], 5)).Any())
                return dims.Where(x => Math.Round(diameter, 5) <= Math.Round(x["size"], 5)).First()["spacing"] / 304.8;
            else return dims.Last()["spacing"] / 304.8;
        }

        int GetSystemRank(string worksetName)
        {
            for (int i = 0; i < AllWorksetNames.Count; i++)
            {
                List<string> worksetNamesList = AllWorksetNames[i];
                if (worksetNamesList.Where(x => x == worksetName).Any()) return i;
            }
            return -1;
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


        private double getRod(XYZ p, FaceArray faces)
        {
            Line tempLine = Line.CreateUnbound(p, XYZ.BasisZ);
            Face lower = faces.get_Item(0);
            Face upper = faces.get_Item(1);
            lower.Intersect(tempLine, out IntersectionResultArray intersectionWithLower);
            //upper.Intersect(tempLine, out IntersectionResultArray intersectionWithUpper);
            if (intersectionWithLower == null || intersectionWithLower.IsEmpty) return 0;
            //if () return 0;
            XYZ ipWithLower = intersectionWithLower.get_Item(0).XYZPoint;
            //XYZ ipWithUpper = intersectionWithUpper.get_Item(0).XYZPoint;
            if (ipWithLower.Z > p.Z)
            {
                return ipWithLower.Z - p.Z;
            }
            else
            {
                return 0;
            }
        }
    }
}
