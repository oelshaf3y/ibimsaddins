using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Structure;
using Line = Autodesk.Revit.DB.Line;
using Parameter = Autodesk.Revit.DB.Parameter;
using System.Windows.Forms;
using Autodesk.Revit.DB.Electrical;
using System.Text;
using System.Drawing;
using System.Windows.Controls;
namespace IBIMSGen.Hangers
{
    [Transaction(TransactionMode.Manual)]
    public class Hangers : IExternalCommand
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
        IList<Element> MechanicalEquipment, ducts, pipes, cables, floors, floooors, ductfits, CTS;
        List<ElementId> plvids;
        IList<Reference> mechRefs, linkedRefs;
        List<Face> floorFacesUp, floorFacesDown;
        List<double> WSdias, WSspcs, CHWdias, CHWspcs, DRdias, DRespcs, Firedias, Firespcs;
        List<double> HangDias1, HangDias2, widthes, pelevsct, floorElevations;
        List<XYZ> HOcts, FOcts;
        List<List<XYZ>> CTPS;
        List<Workset> worksets;
        List<WorksetId> worksetIDs;
        Options options;
        RevitLinkInstance RLI;
        List<DuctHanger> ductHangers;
        List<PipeHanger> pipeHangers;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet element)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            UI = new HangersFM();
            options = new Options();
            options.ComputeReferences = true;
            ductHangers = new List<DuctHanger>();
            pipeHangers = new List<PipeHanger>();

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
            RLI = linksFEC.Cast<RevitLinkInstance>()
                .Where(x => ((RevitLinkType)doc.GetElement(x.GetTypeId())).Name == LinksNames[UI.lnk]).First();
            LinkDoc = RLI.GetLinkDocument();

            AllWorksetNames = UI.AllworksetsNames;
            AllWorksetsDIMS = UI.AllworksetsDIMS;
            //==================================================================================
            double ductOffset = 100 / 304.8;
            double negLength = 100 / 304.80;
            //==================================================================================
            mechRefs = new List<Reference>();
            linkedRefs = new List<Reference>();
            ducts = new List<Element>();
            pipes = new List<Element>();
            cables = new List<Element>();
            floors = new List<Element>();



            try
            {
                mechRefs = uidoc.Selection.PickObjects(ObjectType.Element, new SelectionFilterPDC(), "Select Pipes / Ducts / CableTrays.");
            }
            catch (Exception ex)
            {
                return Result.Cancelled;
            }
            if (UI.selc) // allows me to select region
            {
                try
                {
                    linkedRefs = uidoc.Selection.PickObjects(ObjectType.LinkedElement, "Select Linked Elements.");
                }
                catch (Exception ex)
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
            if (ducts.Count == 0 && pipes.Count == 0 && cables.Count == 0)
            {
                TaskDialog.Show("Wrong Selection", "There are no Ducts or Pipes or Cabletrays selected.");
                return Result.Failed;
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

            floorElevations = new List<double>();
            floooors = new List<Element>();
            floorFacesUp = new List<Face>();
            floorFacesDown = new List<Face>();
            if (floors.Count == 0)
            {
                TaskDialog.Show("Error", "Linked Revit has no Floors to Host Hangers." + Environment.NewLine + "Make sure that Linked Revit is Structural discipline and has Floors.");
                return Result.Failed;
            }

            foreach (Element ele in floors)
            {
                GeometryElement gele = ele.get_Geometry(options);
                foreach (GeometryObject geo in gele)
                {
                    Solid solid = geo as Solid;
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

                            floorFacesUp.Add(solid.Faces.get_Item(1)); //upperface
                            floorFacesDown.Add(solid.Faces.get_Item(0)); //lowerface
                            floooors.Add(ele);
                            floorElevations.Add(elevationAtBot);
                            break;
                        }
                    }
                }
            }
            ductfits = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_DuctFitting)
                .OfClass(typeof(FamilyInstance))
                .Where(x => ((LocationPoint)x.Location).Point.Z <= floorElevations.Max() && ((LocationPoint)x.Location).Point.Z >= floorElevations.Min())
                .ToList();


            //TODO | listAdd modify
            WSdias = ListAdd(1)[0];
            WSspcs = ListAdd(1)[1];
            CHWdias = ListAdd(2)[0];
            CHWspcs = ListAdd(2)[1];
            DRdias = ListAdd(3)[0];
            DRespcs = ListAdd(3)[1];
            Firedias = ListAdd(4)[0];
            Firespcs = ListAdd(4)[1];
            //===============================================================================================

            #region ducts
            foreach (Element duct in ducts)
            {
                double spacingFin = 0;
                Curve ductCurve = ((LocationCurve)duct.Location).Curve;
                List<XYZ> pts = new List<XYZ>();
                double ductWidth = 0;
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

                #region duct fitting
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
                    catch (Exception ex) { }
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
                List<XYZ> hangPts = new List<XYZ>();
                if (fittingPts.Count == 0) // No ductfittings
                {
                    if (ductCurve.Length > negLength && ductCurve.Length <= 4 * ductOffset)
                    {
                        hangPts.Add(ductMidPt);
                    }
                    else if (ductCurve.Length <= spacingFin && ductCurve.Length > 4 * ductOffset)
                    {
                        hangPts.Add(Ps);
                        hangPts.Add(Pe);
                    }
                    else if (ductCurve.Length > spacingFin) // collect point of hangers
                    {
                        hangPts.Add(Ps);
                        double n = Math.Ceiling(cc.Length / spacingFin) - 1;
                        XYZ prevPt = Ps;
                        for (int i = 0; i < n; i++)
                        {
                            XYZ point = prevPt.Add(spacingFin * dir);
                            hangPts.Add(point);
                            prevPt = point;
                        }
                        hangPts.Add(Pe);
                    }
                }
                else       // With ductfittings
                {
                    td("duct fits considered" + ductfits.Count.ToString());
                    if (ductCurve.Length > negLength && ductCurve.Length <= 4 * ductOffset)
                    {
                        List<XYZ> fitPtsOrdered = decOrder(new List<XYZ>() { fittingPts[0], fittingPts[1], ductMidPt }, ductCurve);
                        if (fitPtsOrdered.IndexOf(ductMidPt) == 1)
                        {
                            hangPts.Add(fitPtsOrdered[2]); // farther point
                        }
                        else
                        {
                            hangPts.Add(ductMidPt); // duct center
                        }
                    }
                    else if (ductCurve.Length <= spacingFin && ductCurve.Length > 4 * ductOffset)
                    {
                        List<XYZ> fitPtsOrdered = decOrder(new List<XYZ>() { fittingPts[0], fittingPts[1], Ps }, ductCurve);
                        if (fitPtsOrdered.IndexOf(Ps) == 1)
                        {
                            hangPts.Add(fitPtsOrdered[2]);
                        }
                        else
                        {
                            hangPts.Add(Ps);
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
                                        hangPts.Add(pso[2]);
                                    }
                                    else
                                    {
                                        hangPts.Add(pso[0]);
                                    }
                                    break;
                                }
                            }
                        }
                        hangPts.Add(Pe);
                    }
                    else if (ductCurve.Length > spacingFin)
                    {
                        List<XYZ> ps1o = decOrder(new List<XYZ>() { fittingPts[0], fittingPts[1], Ps }, ductCurve);
                        if (ps1o.IndexOf(Ps) == 1)
                        {
                            hangPts.Add(ps1o[2]);
                        }
                        else
                        {
                            hangPts.Add(Ps);
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
                            hangPts.Add(point);
                            prevPt = point;
                        }
                    }
                }
                ductHangers.Add(new DuctHanger(duct, ductPerpendicular, ductWidth, ductHeight, insoThick, botElevation, hangPts));

            }
            #endregion

            HangDias1 = new List<double>() { 17, 22, 27, 34, 42, 52, 65, 67, 77, 82, 92, 102, 112, 127, 152, 162, 202 };
            HangDias2 = new List<double>() { 227, 252, 317, 352, 402 };

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
                foreach (double hangerDiam in HangDias1)
                {
                    if (pipeDiam <= (hangerDiam / 304.8))
                    {
                        hangerDiameter = hangerDiam / 304.8;
                        break;
                    }
                }
                if (hangerDiameter == -1)
                {
                    foreach (double hangerDiam in HangDias2)
                    {
                        if (pipeDiam <= (hangerDiam / 304.8))
                        {
                            hangerDiameter = hangerDiam / 304.8;
                            break;
                        }
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
                    AddAdd(pipeCurve.Evaluate(0.50, true), pipeHangPts);
                }
                else if (pipeCurve.Length <= spacing && pipeCurve.Length > pipeOffset)
                {
                    AddAdd(Ps, pipeHangPts);
                    AddAdd(Pe, pipeHangPts);
                }
                else if (pipeCurve.Length > spacing)
                {
                    AddAdd(Ps, pipeHangPts);
                    double n = Math.Ceiling(hangCurve.Length / spacing) - 1;
                    XYZ prev = Ps;
                    double Ns = (hangCurve.Length / (n + 1));
                    for (int i = 0; i < n; i++)
                    {
                        XYZ point = prev.Add(Ns * pipeDirection);
                        AddAdd(point, pipeHangPts);
                        prev = point;
                    }
                    AddAdd(Pe, pipeHangPts);
                }

                pipeHangers.Add(new PipeHanger(pipe, P0, isFireFighting, hangerDiameter, levelId, pipeElevation, midElevStart, midElevEnd, slope, pipeHangPts, pipeDirection, pipePerpendicular));
            }
            widthes = new List<double>();
            plvids = new List<ElementId>();
            pelevsct = new List<double>();
            CTPS = new List<List<XYZ>>();
            CTS = new List<Element>();
            HOcts = new List<XYZ>();
            FOcts = new List<XYZ>();
            foreach (Element p in cables)
            {
                Curve c = ((LocationCurve)p.Location).Curve; double Ng = 100 / 304.8; double ff = 500 / 304.80;
                XYZ FOp = ((Line)c).Direction.Normalize(); XYZ HOp = new XYZ(-FOp.Y, FOp.X, FOp.Z);
                XYZ P0 = c.Evaluate(0, true); XYZ Pf = c.Evaluate(1, true);
                XYZ Ps = P0.Add(ductOffset * FOp); XYZ Pe = Pf.Add(-ductOffset * FOp);
                double s = 1500 / 304.8;
                Curve cc = null;
                try
                {
                    cc = Line.CreateBound(Ps, Pe) as Curve;
                }
                catch { continue; }
                List<XYZ> pps = new List<XYZ>(); List<XYZ> pangpts = new List<XYZ>();
                double width = p.LookupParameter("Width").AsDouble(); widthes.Add(width);
                ElementId plvlid = p.LookupParameter("Reference Level").AsElementId(); plvids.Add(plvlid);
                double BE = p.LookupParameter("Bottom Elevation").AsDouble();
                double pelev = ((Level)doc.GetElement(plvlid)).Elevation + BE; pelevsct.Add(pelev);
                if (c.Length > Ng && c.Length <= ff)
                {
                    AddAdd(c.Evaluate(0.50, true), pangpts);
                }
                else if (c.Length <= s && c.Length > ff)
                {
                    AddAdd(Ps, pangpts);
                    AddAdd(Pe, pangpts);
                }
                else if (c.Length > s)
                {
                    AddAdd(Ps, pangpts);
                    double n = Math.Floor((cc.Length + (100 / 304.8)) / s);
                    XYZ Pis = Ps;
                    for (int i = 0; i < n; i++)
                    {
                        XYZ Pie = Pis.Add(s * FOp);
                        AddAdd(Pie, pangpts);
                        Pis = Pie;
                    }
                }
                CTPS.Add(pangpts); CTS.Add(p);
                HOcts.Add(FOp); FOcts.Add(HOp);
            }
            string err = ""; int errco = 0;
            //====================================================================================================
            //====================================================================================================
            using (Transaction trans = new Transaction(doc, "IBIMS_Hangers"))
            {
                trans.Start();
                int h = 0;
                /*fsd2.Activate();*/
                pipeHanger20.Activate(); ductHanger.Activate(); pipeHanger2.Activate(); pipeHanger200.Activate();
                foreach (DuctHanger hanger in ductHangers)
                {

                    if (hanger.hangPts.Count > 0)
                    {
                        foreach (XYZ p in hanger.hangPts)
                        {
                            FamilySymbol FS = null;
                            int fflu = -1;
                            int ffld = -1;
                            double ROD = 0;
                            double dd = double.MinValue;
                            double dd2 = double.MaxValue;
                            double Zu = 0; double Zd = 0;
                            foreach (Face face in floorFacesDown)
                            {
                                Line ll = Line.CreateUnbound(p, XYZ.BasisZ);
                                IntersectionResultArray iraa = new IntersectionResultArray();
                                SetComparisonResult scr = face.Intersect(ll, out iraa);
                                if (iraa != null)
                                {
                                    if (!iraa.IsEmpty)
                                    {
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
                            if (fflu != -1 && ffld != -1)
                            {
                                if (Zu <= Zd)
                                {
                                    ROD = Zu;
                                    //FS = fsd1;
                                }
                                else
                                {
                                    ROD = Zd;
                                    //FS = fsd2;
                                }
                            }
                            else if (fflu != -1)
                            {
                                ROD = Zu;
                                //FS = fsd1;
                            }
                            else if (ffld != -1)
                            {
                                ROD = Zd;
                                //FS = fsd2;
                            }
                            else
                            {
                                ROD = 2000 / 304.8;
                                //FS = fsd1;
                            }
                            XYZ PP = new XYZ(p.X, p.Y, p.Z - hanger.insoThick - (hanger.ductHeight / 2));
                            FamilyInstance hang = doc.Create.NewFamilyInstance(PP, ductHanger, hanger.ductPerpendicular, hanger.duct, StructuralType.NonStructural);
                            hang.LookupParameter("Width").Set(hanger.ductWidth + (2 * hanger.insoThick) + 16 / 304.8);
                            //if (FS == fsd1)
                            //{
                            double Z = hanger.botElevation - hanger.insoThick - hang.LookupParameter("Elevation from Level").AsDouble();
                            hang.Location.Move(new XYZ(0, 0, Z));
                            ROD += hanger.insoThick + (hanger.ductHeight / 2);
                            hang.LookupParameter("ROD 1").Set(ROD);
                            hang.LookupParameter("ROD 2").Set(ROD);
                            //}
                            //else if (FS == fsd2)+ (124 / 304.8)
                            //{
                            //    double fltop = floorelevs[ffld];
                            //    double Z = fltop - hang.LookupParameter("Elevation from Level").AsDouble();
                            //    hang.Location.Move(new XYZ(0, 0, Z));
                            //    hang.LookupParameter("Height").Set(Belevs[h] - insthicks[h] - fltop);
                            //}
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
                int ctco = 0;
                foreach (Element ct in CTS)
                {
                    foreach (XYZ p in CTPS[ctco])
                    {
                        int fflu = -1; double ddu = double.MinValue;
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
                            XYZ PP = new XYZ(p.X, p.Y, pelevsct[ctco]); ductHanger.Activate();
                            FamilyInstance hang = doc.Create.NewFamilyInstance(PP, ductHanger, FOcts[ctco], doc.GetElement(plvids[ctco]), StructuralType.NonStructural);
                            hang.LookupParameter("Width").Set(widthes[ctco] + 100 / 304.8);
                            hang.LookupParameter("ROD 1").Set(ROD);
                            hang.LookupParameter("ROD 2").Set(ROD);
                        }
                    }
                    ctco++;
                }
                if (errco > 0)
                {
                    TaskDialog.Show("Warning", errco + " Pipes have not Hangers." + Environment.NewLine + err);
                }
                trans.Commit();
            }
            return Result.Succeeded;
        }
        void AddAdd(XYZ p, List<XYZ> pss)
        {
            if (!pss.Contains(p))
            {
                pss.Add(p);
            }
        }
        Face GetFaces(Element ele)
        {
            Face face = null;
            options.View = doc.ActiveView;
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
        List<List<double>> ListAdd(int R)
        {
            List<List<double>> list = new List<List<double>>();
            List<double> dias = new List<double>();
            List<double> spacs = new List<double>();
            int coun = 0;
            foreach (string ss in AllWorksetsDIMS[R][0])
            {
                double dia = Convert.ToDouble(ss); double spac = Convert.ToDouble(AllWorksetsDIMS[R][1][coun]);
                if (dia != 0 && spac != 0)
                {
                    dias.Add(dia); spacs.Add(spac / 304.8);
                }
                coun++;
            }
            list.Add(dias); list.Add(spacs);
            return list;
        }
        double SysSpacing(int rankk, double dia)
        {
            double spac = 0;
            if (rankk == 1)
            {
                spac = GetSpace(WSdias, WSspcs, dia);
            }
            else if (rankk == 2)
            {
                spac = GetSpace(CHWdias, CHWspcs, dia);
            }
            else if (rankk == 3)
            {
                spac = GetSpace(DRdias, DRespcs, dia);
            }
            else if (rankk == 4)
            {
                spac = GetSpace(Firedias, Firespcs, dia);
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
        void td(string message)
        {
            TaskDialog.Show("Info", message);
        }
        List<XYZ> decOrder(List<XYZ> oldlist, Curve cu)
        {
            List<XYZ> newlist = new List<XYZ>();
            if (Math.Round(((Line)cu).Direction.Normalize().Y, 3) == 0)
            {
                newlist = oldlist.OrderByDescending(a => a.X).ToList();
            }
            else
            {
                newlist = oldlist.OrderByDescending(a => a.Y).ToList();
            }
            return newlist;
        }
    }
}
