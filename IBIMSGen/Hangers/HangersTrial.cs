using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
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
        List<string> LinksNames, levelsNames, worksetnames;
        List<FamilySymbol> familySymbols;
        Document LinkDoc;
        List<Level> levels;
        List<List<string>> AllWorksetNames;
        List<List<Dictionary<string, double>>> AllWorksetsDIMS;
        IList<Element> ducts, pipes, cables, floors;
        IList<Reference> mechRefs, linkedRefs;
        List<Element> ductfits;
        List<double> HangDias;
        List<Workset> worksets;
        List<WorksetId> worksetIDs;
        Options options;
        RevitLinkInstance RLI;
        List<RevitLinkInstance> ductsRLI, pipesRLI;
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
            ductsRLI = new List<RevitLinkInstance>();
            pipesRLI = new List<RevitLinkInstance>();
            ductHangers = new List<DuctHanger>();
            pipeHangers = new List<PipeHanger>();
            trayHangers = new List<TrayHanger>();
            mechRefs = new List<Reference>();
            linkedRefs = new List<Reference>();
            ducts = new List<Element>();
            pipes = new List<Element>();
            cables = new List<Element>();
            floors = new List<Element>();
            ductfits = new List<Element>();
            ductOffset = 100 / 304.8;
            negLength = 100 / 304.80;
            HangDias = new List<double>() { 17, 22, 27, 34, 42, 52, 65, 67, 77, 82, 92, 102, 112, 127, 152, 162, 202, 227, 252, 317, 352, 402 };

            familySymbols = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
            //pipeHanger20 = familySymbols.Where(x => x.FamilyName.Equals("02- PIPE HANGER ( 20 - 200 )"))?.FirstOrDefault() ?? null;
            ////pipeHanger200 = familySymbols.Where(x => x.FamilyName.Equals("01- PIPE HANGER ( +200 mm )"))?.FirstOrDefault() ?? null;
            //pipeHanger2 = familySymbols.Where(x => x.FamilyName.Equals("Pipes Hanger 2"))?.FirstOrDefault() ?? null;
            //ductHanger = familySymbols.Where(x => x.FamilyName.Equals("The Lower Bridge Duct Hanger"))?.FirstOrDefault() ?? null;

            StringBuilder nullfams = new StringBuilder();
            //if (pipeHanger20 == null) { nullfams.AppendLine("02- PIPE HANGER ( 20 - 200 )"); }
            //if (pipeHanger200 == null) { nullfams.AppendLine("01- PIPE HANGER ( +200 mm )"); }
            //if (pipeHanger2 == null) { nullfams.AppendLine("Pipes Hanger 2"); }
            //if (ductHanger == null) { nullfams.AppendLine("The Lower Bridge Duct Hanger"); }
            //if (ductHanger == null || pipeHanger20 == null || pipeHanger200 == null || pipeHanger2 == null)
            //{
            //    TaskDialog.Show("Error", "Please Load Supports Family.\n" + nullfams.ToString());
            //    return Result.Failed;
            //}
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
            UI = new HangersFM(LinksNames, levelsNames, familySymbols);
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
                if (UI.useLink.Checked)
                {
                    mechRefs = uidoc.Selection.PickObjects(ObjectType.LinkedElement, new SelectionFilterPDC(RLI), "Select Pipes / Ducts / CableTrays.");
                    //after selection process
                    foreach (Reference reference in mechRefs)
                    {
                        Element elem = doc.GetElement(reference.ElementId);

                        RevitLinkInstance rli = elem as RevitLinkInstance;
                        Document _link = rli.GetLinkDocument();
                        Element e = _link.GetElement(reference.LinkedElementId);
                        if (e.Category.Name == "Ducts")
                        {
                            Curve ductCurve = ((LocationCurve)e.Location).Curve;
                            double s1 = Math.Round(e.LookupParameter("Start Middle Elevation").AsDouble(), 6);
                            double s2 = Math.Round(e.LookupParameter("End Middle Elevation").AsDouble(), 6);
                            if (ductCurve.Length >= negLength && s1 == s2) //perfectly horizontal 
                            {
                                ducts.Add(e);
                                ductsRLI.Add(rli);
                            }
                        }
                        else if (e.Category.Name == "Pipes")
                        {
                            Curve pipeCurve = ((LocationCurve)e.Location).Curve;
                            if (pipeCurve.Length >= negLength && Math.Abs(Math.Round(((Line)pipeCurve).Direction.Normalize().Z, 3)) != 1)
                            {
                                pipes.Add(e);
                                pipesRLI.Add(rli);
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
                else
                {
                    mechRefs = uidoc.Selection.PickObjects(ObjectType.Element, new SelectionFilterPDC(), "Select Pipes / Ducts / CableTrays.");


                    //after selection process
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

            }
            catch (Exception ex)
            {
                td(ex.StackTrace);
                //return Result.Cancelled;
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
            List<RevitLinkInstance> fitsRLI = new List<RevitLinkInstance>();
            foreach (RevitLinkInstance rli in ductsRLI)
            {
                if (fitsRLI.Where(x => x.Name == rli.Name).Any()) continue;
                fitsRLI.Add(rli);
            }
            if (ductsRLI.Count > 0)
            {
                foreach (RevitLinkInstance rli in fitsRLI)
                {
                    Document link = rli.GetLinkDocument();
                    ductfits.AddRange(new FilteredElementCollector(link)
                        .OfCategory(BuiltInCategory.OST_DuctFitting)
                        .OfClass(typeof(FamilyInstance))
                        //.Where(x => selectionBoundary.contains(x))
                        );
                }
            }
            else
            {

                ductfits = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_DuctFitting)
                    .OfClass(typeof(FamilyInstance))
                    //.Where(x => selectionBoundary.contains(x))
                    .ToList();
            }

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
                        List<Element> fitsInRange = fitsTree.query(floorRange);
                        List<Element> pipesInRange = pipesTree.query(floorRange);
                        List<Element> traysInRange = traysTree.query(floorRange);

                        watch.Stop();
                        sb.AppendLine($"floor loop {watch.ElapsedMilliseconds}");
                        watch.Restart();
                        foreach (Element duct in ductTree.query(floorRange))
                        {
                            ductHangers.Add(new DuctHanger(doc, solid, duct, AllWorksetsDIMS, familySymbols, negLength, ductOffset, floorUp, floorDown, fitsInRange));
                        }
                        watch.Stop();
                        sb.AppendLine($"Proccess ducts end {watch.ElapsedMilliseconds}");
                        watch.Restart();
                        #region Pipes
                        for (int i = 0; i < pipesInRange.Count; i++)
                        {
                            Element pipe = pipesInRange[i];
                            WorksetId wsid = pipe.WorksetId;
                            string Workset = worksetnames[worksetIDs.IndexOf(wsid)];
                            double PipeOffset = 500 / 304.80;
                            pipeHangers.Add(new PipeHanger(doc, solid, pipe, floorUp, floorDown, AllWorksetsDIMS,
                                familySymbols, negLength, PipeOffset, Workset, HangDias, AllWorksetNames, RLI, pipesRLI[i]));
                        }
                        #endregion
                        //processPipes(ele as Floor, pipesInRange, solid.Faces, floorUp, floorDown, elevationAtBot);
                        watch.Stop();
                        sb.AppendLine($"Proccess Pipes end {watch.ElapsedMilliseconds}");
                        watch.Restart();
                        //processTrays(traysInRange, solid.Faces, floorUp, floorDown);
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
                foreach (DuctHanger ductHanger in ductHangers)
                {
                    ductHanger.Plant();
                }
                #region pipes
                foreach (PipeHanger pipeHanger in pipeHangers)
                {
                    pipeHanger.Plant();
                }
                #endregion

                #region trays
                foreach (TrayHanger tray in trayHangers)
                {
                    foreach (Support support in tray.Supports)
                    {
                        tray.FamilySymbol.Activate();
                        FamilyInstance hang = doc.Create.NewFamilyInstance(support.point, tray.FamilySymbol, tray.Perpendicular, doc.GetElement(tray.levelId), StructuralType.NonStructural);
                        hang.LookupParameter("Width").Set(tray.Width + 100 / 304.8);
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




        public void processTrays(List<Element> traysInRange, FaceArray faces, double floorUp, double floorDown)
        {
            #region Cable Trays
            foreach (Element tray in traysInRange)
            {
                FamilySymbol trayFamilySymbol = null;
                Curve trayCurve = ((LocationCurve)tray.Location).Curve;
                double trayOffset = 500 / 304.80;
                XYZ trayDir = ((Line)trayCurve).Direction.Normalize();
                XYZ perpendicular = new XYZ(-trayDir.Y, trayDir.X, trayDir.Z);
                XYZ P0 = trayCurve.Evaluate(0, true);
                XYZ Pf = trayCurve.Evaluate(1, true);
                XYZ Ps = P0.Add(ductOffset * trayDir);
                XYZ Pe = Pf.Add(-ductOffset * trayDir);
                double width = tray.LookupParameter("Width").AsDouble();
                double spacingFin = 0;
                var a = AllWorksetsDIMS[5][0];
                if (AllWorksetsDIMS[5].Count == 1)
                {
                    spacingFin = AllWorksetsDIMS[5][0]["spacing"] / 304.8;
                    if (spacingFin == 0) continue;
                    string familySymbolName = familySymbols.Select(x => x.FamilyName).Distinct().ElementAt(Convert.ToInt32(AllWorksetsDIMS[5][0]["family"]));
                    trayFamilySymbol = familySymbols.Where(x => x.FamilyName.Equals(familySymbolName)).First();
                }
                else
                {
                    if (AllWorksetsDIMS[5].Where(x => width * 304.8 <= x["size"]).Any())
                    {
                        spacingFin = AllWorksetsDIMS[5].Where(x => width * 304.8 <= x["size"]).First()["spacing"] / 304.8;
                        if (spacingFin == 0) continue;
                        string familySymbolName = familySymbols.Select(x => x.FamilyName).Distinct().ElementAt(Convert.ToInt32(AllWorksetsDIMS[5][0]["family"]));
                        trayFamilySymbol = familySymbols.Where(x => x.FamilyName.Equals(familySymbolName)).First();
                    }
                }
                Curve hangCurve = null;
                try
                {
                    hangCurve = Line.CreateBound(Ps, Pe);
                }
                catch { continue; }
                List<XYZ> trayHangPts = new List<XYZ>();
                List<Support> supports = new List<Support>();
                ElementId levelId = tray.LookupParameter("Reference Level").AsElementId();
                double botElevation = tray.LookupParameter("Lower End Bottom Elevation").AsDouble();
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
                else if (trayCurve.Length <= spacingFin && trayCurve.Length > trayOffset)
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
                else if (trayCurve.Length > spacingFin)
                {
                    if (!trayHangPts.Contains(Ps))
                    {
                        trayHangPts.Add(Ps);
                        double rod = getRod(Ps, faces);
                        XYZ PP = new XYZ(Ps.X, Ps.Y, elevation);
                        if (rod != 0) supports.Add(new Support(PP, rod));
                    }
                    double n = Math.Floor((hangCurve.Length + (100 / 304.8)) / spacingFin);
                    XYZ prevPt = Ps;
                    for (int i = 0; i < n; i++)
                    {
                        XYZ point = prevPt.Add(spacingFin * trayDir);
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

                trayHangers.Add(new TrayHanger(tray, width, levelId, supports, trayDir, perpendicular, trayFamilySymbol));
            }
            #endregion
        }

        //interface implemented
        double SysSpacing(List<Dictionary<string, double>> dims, double diameter)
        {
            if (dims == null) return 0;
            if (dims.Count == 0) return 0;
            else if (dims.Count == 1) return dims[0]["spacing"];
            else if (dims.Where(x => Math.Round(diameter, 5) <= Math.Round(x["size"], 5)).Any())
                return dims.Where(x => Math.Round(diameter, 5) <= Math.Round(x["size"], 5)).First()["spacing"] / 304.8;
            else return dims.Last()["spacing"] / 304.8;
        }

        //interface implemented
        int GetSystemRank(string worksetName)
        {
            for (int i = 0; i < AllWorksetNames.Count; i++)
            {
                List<string> worksetNamesList = AllWorksetNames[i];
                if (worksetNamesList.Where(x => x == worksetName).Any()) return i;
            }
            return -1;
        }

        //interface implemented
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
                    return (Solid)geo.FirstOrDefault();
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
                return solids.OrderByDescending(x => x.Volume).ElementAt(0);
            }
            else
            {
                return null;
            }
        }

        //interface implemented
        private double getRod(XYZ point, FaceArray faces)
        {
            Line tempLine = Line.CreateUnbound(point, XYZ.BasisZ);
            Face lower = faces.get_Item(0);
            Face upper = faces.get_Item(1);
            lower.Intersect(tempLine, out IntersectionResultArray intersectionWithLower);
            //upper.Intersect(tempLine, out IntersectionResultArray intersectionWithUpper);
            if (intersectionWithLower == null || intersectionWithLower.IsEmpty) return 0;
            //if () return 0;
            XYZ ipWithLower = intersectionWithLower.get_Item(0).XYZPoint;
            //XYZ ipWithUpper = intersectionWithUpper.get_Item(0).XYZPoint;
            if (ipWithLower.Z > point.Z)
            {
                return ipWithLower.Z - point.Z;
            }
            else
            {
                return 0;
            }
        }
    }
}
