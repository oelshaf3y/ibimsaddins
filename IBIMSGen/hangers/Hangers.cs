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
    public class Hangers : IExternalCommand
    {
        UIDocument uidoc;
        Document doc;
        HangersFM UI;
        FilteredElementCollector linksFEC, levelsFEC;
        List<string> LinksNames, levelsNames;
        List<FamilySymbol> familySymbols;
        Document LinkDoc;
        List<Level> levels;
        List<List<string>> AllWorksetNames;
        List<List<Dictionary<string, double>>> AllWorksetsDIMS;
        IList<Element> ducts, pipes, cables, floors;
        IList<Reference> mechRefs, linkedRefs;
        List<Element> ductfits;
        List<double> HangDias;
        Options options;
        RevitLinkInstance RLI;
        List<RevitLinkInstance> ductsRLI, pipesRLI, cablesRLI;
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
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            options = new Options();
            options.ComputeReferences = true;
            options.View = doc.ActiveView;
            ductsRLI = new List<RevitLinkInstance>();
            pipesRLI = new List<RevitLinkInstance>();
            cablesRLI = new List<RevitLinkInstance>();
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

            familySymbols = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>().OrderBy(x=>x.FamilyName).ToList();
            linksFEC = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance));
            LinksNames = linksFEC.Cast<RevitLinkInstance>()
                .Select(x => ((RevitLinkType)doc.GetElement(x.GetTypeId())))
                .Where(x => x.GetLinkedFileStatus() == LinkedFileStatus.Loaded).OrderBy(x => x.Name)
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


            if (UI.useLink.Checked)
            {
                try
                {
                    mechRefs = uidoc.Selection.PickObjects(ObjectType.LinkedElement, new SelectionFilterPDC(RLI), "Select Pipes / Ducts / CableTrays.");
                }
                catch
                {
                    return Result.Cancelled;
                }//after selection process
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
                            cablesRLI.Add(rli);
                        }
                    }

                }
            }
            else
            {
                try
                {
                    mechRefs = uidoc.Selection.PickObjects(ObjectType.Element, new SelectionFilterPDC(), "Select Pipes / Ducts / CableTrays.");
                }
                catch
                {
                    return Result.Cancelled;
                }

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


            if (UI.selc) // allows me to select region
            {
                try
                {
                    linkedRefs = uidoc.Selection.PickObjects(ObjectType.LinkedElement, new SelectionFilterPDC(isHost: true), "Select Linked Elements.");

                }
                catch
                {
                    return Result.Cancelled;
                }
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
            double count = 0;
            foreach (Element tray in cables)
            {
                if (traysTree.insert(tray)) count++;
            }

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
                        foreach (Element duct in ductTree.query(floorRange))
                        {
                            ductHangers.Add(new DuctHanger(doc, solid, duct, AllWorksetsDIMS, familySymbols, negLength, ductOffset, floorUp, floorDown, fitsInRange));
                            //new DuctHanger(doc, solid, duct, AllWorksetsDIMS, familySymbols, negLength, ductOffset, floorUp, floorDown, fitsInRange);
                        }
                        #region Pipes
                        for (int i = 0; i < pipesInRange.Count; i++)
                        {
                            Element pipe = pipesInRange[i];
                            if (UI.useLink.Checked) pipeHangers.Add(new PipeHanger(doc, solid, pipe, floorUp, floorDown, AllWorksetsDIMS,
                                familySymbols, negLength, offset: 500 / 304.80, HangDias, AllWorksetNames, RLI, pipesRLI[i]));
                            else pipeHangers.Add(new PipeHanger(doc, solid, pipe, floorUp, floorDown, AllWorksetsDIMS,
                                familySymbols, negLength, offset: 500 / 304.80, HangDias, AllWorksetNames, RLI));
                            //new PipeHanger(doc, solid, pipe, floorUp, floorDown, AllWorksetsDIMS,
                            //    familySymbols, negLength, offset: 500 / 304.80, HangDias, AllWorksetNames, RLI, pipesRLI[i]);
                        }
                        #endregion
                        //processPipes(ele as Floor, pipesInRange, solid.Faces, floorUp, floorDown, elevationAtBot);

                        #region Cable Trays
                        for (int i = 0; i < traysInRange.Count; i++)
                        {
                            Element tray = traysInRange[i];
                            if (UI.useLink.Checked) trayHangers.Add(new TrayHanger(doc, solid, tray, AllWorksetsDIMS, floorUp, floorDown, familySymbols, negLength, ductOffset, RLI, cablesRLI[i]));
                            else trayHangers.Add(new TrayHanger(doc, solid, tray, AllWorksetsDIMS, floorUp, floorDown, familySymbols, negLength, ductOffset, RLI));
                        }
                        #endregion
                    }
                }
            }

            //====================================================================================================
            //====================================================================================================
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
                foreach (DuctHanger ductHanger in ductHangers)
                {
                    ductHanger.Plant();
                }
                foreach (PipeHanger pipeHanger in pipeHangers)
                {
                    pipeHanger.Plant();
                }
                #region trays
                foreach (TrayHanger tray in trayHangers)
                {
                    tray.Plant();
                }
                #endregion

                tr.Commit();
                tr.Dispose();
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

    }
}
