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
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using System.Text;
using System.Drawing;

namespace IBIMS_MEP
{
    [Transaction(TransactionMode.Manual)]
    public class Hangers : IExternalCommand
    {
        UIDocument uidoc;
        Document doc, linkDoc;
        Autodesk.Revit.Creation.Application app;
        HangersFM uiForm;

        StringBuilder sb;

        FilteredElementCollector genModelFamInstFEC, levels, pipeFitFamInstFEC, linkedFloorsFEC, ductFitingFamInsFEC;
        ElementCategoryFilter elemCatFilterGenMods, elemCatFilterPipeFit, elemCatFilterDuctFit;

        IList<Element> MechanicalEquipment, floors, cables, pipes, ducts, pipefits, ductFits;
        IList<Reference> linkedElemsRefs, refsToHang;

        FamilySymbol pipeHanger20, pipeHanger2, pipeHanger200, ductHanger;

        List<string> LinksNames, lvlNames, worksetnames;
        List<List<string>> AllWorksetNames;
        List<List<List<string>>> AllWorksetsDIMS;

        List<double> elevations, floorElevations;
        List<double> fireSpacers, fireDiams, drainageSpacers, drainageDiams, chillWaterSpacers, chillWaterDiams, waterSuplySpacers, waterSuplyDiams;
        List<Level> lvls;
        List<Workset> worksets;
        List<WorksetId> worksetIDs;
        List<Face> floorfaces, floorfacesd;

        RevitLinkInstance RLI;
        Level toLevel, fromLevel;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet element)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = commandData.Application.ActiveUIDocument.Document;
            app = commandData.Application.Application.Create;

            refsToHang = new List<Reference>();
            linkedElemsRefs = new List<Reference>();
            ducts = new List<Element>();
            pipes = new List<Element>();
            cables = new List<Element>();
            floors = new List<Element>();

            genModelFamInstFEC = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol));
            elemCatFilterGenMods = new ElementCategoryFilter(BuiltInCategory.OST_GenericModel);

            MechanicalEquipment = genModelFamInstFEC.WherePasses(elemCatFilterGenMods).ToList();
            pipeHanger20 = null; pipeHanger2 = null; pipeHanger200 = null; ductHanger = null;
            sb = new StringBuilder();

            #region Getting Family Symbols
            foreach (Element elem in MechanicalEquipment)
            {
                FamilySymbol fi = elem as FamilySymbol;
                if (fi.FamilyName == "02- PIPE HANGER ( 20 - 200 )")
                {
                    pipeHanger20 = fi;
                }
                else if (fi.FamilyName == "01- PIPE HANGER ( +200 mm )")
                {
                    pipeHanger200 = fi;
                }
                else if (fi.FamilyName == "Pipes Hanger 2")
                {
                    pipeHanger2 = fi;
                }
                else if (fi.FamilyName == "The Lower Bridge Duct Hanger")
                {
                    ductHanger = fi;
                }
                if (pipeHanger20 != null && pipeHanger2 != null && ductHanger != null && pipeHanger200 != null) { break; } //break when all family symbols are collected
            }
            if (pipeHanger20 == null)
            {
                sb.AppendLine("02- PIPE HANGER ( 20 - 200 )");
            }
            if (pipeHanger200 == null)
            {
                sb.AppendLine("01- PIPE HANGER ( +200 mm )");
            }
            if (pipeHanger2 == null)
            {
                sb.AppendLine("Pipes Hanger 2");
            }
            if (ductHanger == null)
            {
                sb.AppendLine("The Lower Bridge Duct Hanger");
            }

            //FilteredElementCollector fec = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol));
            //ElementCategoryFilter ecf = new ElementCategoryFilter(BuiltInCategory.OST_DuctAccessory);
            //IList<Element> generics = fec.WherePasses(ecf).ToList();  /*FamilySymbol fsd2 = null;*/
            //foreach (Element e in generics)
            //{
            //    FamilySymbol fi = e as FamilySymbol;
            //    if (fi.FamilyName == "The Lower Bridge Duct Hanger")
            //    {
            //        fsd1 = fi;
            //        break;
            //    }
            //}
            //foreach (Element e in generics)
            //{
            //    FamilySymbol fi = e as FamilySymbol;
            //    if (fi.FamilyName == "The Upper Bridge Duct Hanger")
            //    {
            //        fsd2 = fi;
            //        break;
            //    }
            //}

            if (ductHanger == null || pipeHanger20 == null || pipeHanger200 == null || pipeHanger2 == null)
            {
                TaskDialog.Show("Error", "Please Load Supports Family." + "\n" + sb.ToString());
                return Result.Failed;
            }
            #endregion

            #region Getting Links

            LinksNames = new List<string>();
            linkDoc = null;
            FilteredElementCollector linksFEC = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance));
            foreach (RevitLinkInstance linkInstance in linksFEC)
            {
                RevitLinkType revitLinkType = (RevitLinkType)doc.GetElement(linkInstance.GetTypeId());
                if (revitLinkType.GetLinkedFileStatus() == LinkedFileStatus.Loaded)
                {
                    string name = revitLinkType.Name;
                    LinksNames.Add(name);
                }
            }
            if (LinksNames.Count == 0)
            {
                TaskDialog.Show("Error", "There are no Loaded Linked Revit detected in Project.");
                return Result.Failed;
            }

            levels = new FilteredElementCollector(doc).OfClass(typeof(Level));
            lvlNames = new List<string>();
            elevations = new List<double>();
            lvls = new List<Level>();

            foreach (Level lvl in levels.OrderBy(x => (x as Level).Elevation))
            {
                elevations.Add(lvl.Elevation);
                lvls.Add(lvl);
                lvlNames.Add(lvl.Name);
            }

            //foreach (Level lvl in levels)
            //{
            //    elevations.Add(lvl.Elevation);
            //}

            //elevations.Sort();
            //foreach (Level lvl in levels)
            //{
            //    double ee = lvl.Elevation;
            //    foreach (double e in elevations)
            //    {
            //        if (ee == e)
            //        {
            //            lvls.Add(lvl);
            //            lvlNames.Add(lvl.Name);
            //            break;
            //        }
            //    }
            //}


            #endregion

            #region Getting Worksets
            FilteredWorksetCollector worksetsFEC = new FilteredWorksetCollector(doc);
            worksets = new List<Workset>();
            worksetIDs = new List<WorksetId>();
            worksetnames = new List<string>();
            foreach (Workset ws in worksetsFEC)
            {
                if (ws.Kind == WorksetKind.UserWorkset)
                {
                    worksetIDs.Add(ws.Id);
                    worksetnames.Add(ws.Name);
                    worksets.Add(ws);
                }
            }
            if (worksetnames.Count == 0)
            {
                TaskDialog.Show("Error", "Document has not UserWorksets.");
                return Result.Failed;
            }

            #endregion

            uiForm = new HangersFM();
            uiForm.Linkes = LinksNames;
            uiForm.worksetnames = worksetnames;
            uiForm.Levels = lvlNames;
            uiForm.ShowDialog();

            if (uiForm.DialogResult == DialogResult.Cancel)
            {
                return Result.Cancelled;
            }

            //getting linked document from the uiform
            #region Getting Linked Document

            RLI = null;
            foreach (RevitLinkInstance ri in linksFEC)
            {
                RevitLinkType revitLinkType = (RevitLinkType)doc.GetElement(ri.GetTypeId());
                LinksNames.Add(revitLinkType.Name);
                if (revitLinkType.Name == LinksNames[uiForm.lnk])
                {
                    linkDoc = ri.GetLinkDocument();
                    RLI = ri;
                    if (linkDoc != null)
                    {
                        break;
                    }
                }
            }
            #endregion

            AllWorksetNames = uiForm.AllworksetsNames;
            AllWorksetsDIMS = uiForm.AllworksetsDIMS;

            //==================================================================================
            double f = 100 / 304.8;
            double Ngd = 100 / 304.80;
            //==================================================================================
            #region Select Elements To Hang
            pipeFitFamInstFEC = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance));
            elemCatFilterPipeFit = new ElementCategoryFilter(BuiltInCategory.OST_PipeFitting);
            pipefits = pipeFitFamInstFEC.WherePasses(elemCatFilterPipeFit).ToList();

            selcFilterPDCt selcFilter = new selcFilterPDCt();
            try
            {
                TaskDialog.Show("Select", "Select Pipes / Ducts / CableTrays.");
                refsToHang = uidoc.Selection.PickObjects(ObjectType.Element, selcFilter);
            }
            catch (Exception ex)
            {
                td(ex.Message);
                return Result.Cancelled;
            }
            #endregion


            floorElevations = new List<double>();
            floorfaces = new List<Face>();
            floorfacesd = new List<Face>();


            #region Getting linked Floors

            if (uiForm.selc)
            {
                TaskDialog.Show("Select", "Select Linked Elements.");
                try
                {
                    linkedElemsRefs = uidoc.Selection.PickObjects(ObjectType.LinkedElement);

                    #region Getting All Linked Floors By Selection
                    foreach (Reference reff in linkedElemsRefs)
                    {
                        RevitLinkInstance rli = doc.GetElement(reff.ElementId) as RevitLinkInstance;
                        if (rli != null)
                        {
                            if (RLI.Name.Trim() == rli.Name.Trim())
                            {
                                Autodesk.Revit.DB.Document doca = rli.GetLinkDocument();
                                Element ele = doca.GetElement(reff.LinkedElementId);
                                if (ele is Floor)
                                {
                                    Tuple<bool, Solid, double> process = processFloor(ele);
                                    if (process.Item1)
                                    {
                                        floorfaces.Add(process.Item2.Faces.get_Item(1)); //1
                                        floorfacesd.Add(process.Item2.Faces.get_Item(0)); //-1
                                        floors.Add(ele);
                                        floorElevations.Add(process.Item3);
                                    }
                                }
                            }
                        }
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    td(ex.Message);
                    return Result.Cancelled;
                }
            }
            else
            {
                #region Getting Linked Floors From Links Automatically.
                fromLevel = lvls[uiForm.frin];
                toLevel = lvls[uiForm.toin];
                linkedFloorsFEC = new FilteredElementCollector(linkDoc).OfClass(typeof(Floor));

                foreach (Floor floor in linkedFloorsFEC)
                {
                    double elevAtBot = 0, elevAtTop = 0;
                    Parameter atBotParam = floor.LookupParameter("Elevation at Bottom");
                    Parameter atTopParam = floor.LookupParameter("Elevation at Top");
                    if (atBotParam != null && atTopParam != null)
                    {
                        elevAtBot = atBotParam.AsDouble();
                        elevAtTop = atTopParam.AsDouble();
                        if (elevAtBot == 0 && elevAtTop == 0)
                        {
                            Parameter levelParam = floor.LookupParameter("Level");
                            if (levelParam != null)
                            {
                                double levelElevation = lvls.Where(x => x.Name.Trim() == levelParam.AsValueString().Trim()).FirstOrDefault().Elevation;
                                double heightOffset = floor.LookupParameter("Height Offset From Level").AsDouble();
                                elevAtBot = levelElevation + heightOffset;
                                if (elevAtBot <= toLevel.Elevation && elevAtBot >= fromLevel.Elevation)
                                {
                                    Tuple<bool, Solid, double> process = processFloor(floor);
                                    if (process.Item1)
                                    {
                                        floorfaces.Add(process.Item2.Faces.get_Item(1)); //1
                                        floorfacesd.Add(process.Item2.Faces.get_Item(0)); //-1
                                        floors.Add(floor);
                                        floorElevations.Add(process.Item3);
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (elevAtBot <= toLevel.Elevation && elevAtBot >= fromLevel.Elevation)
                            {
                                floors.Add(floor);
                            }
                        }
                    }
                }
                #endregion
            }
            #endregion

            #region Organizing Elements Into Worksets
            foreach (Reference reff in refsToHang)
            {
                WorksetId wsid;
                string wsName;
                Curve c;
                Element elem = doc.GetElement(reff);
                if (elem.Category.Name == "Ducts")
                {
                    wsid = elem.WorksetId;
                    wsName = worksetnames[worksetIDs.IndexOf(wsid)];
                    if (GetSystemRank(wsName) != 0) { continue; }
                    else
                    {
                        c = ((LocationCurve)elem.Location).Curve;
                        double s1 = Math.Round(elem.LookupParameter("Start Middle Elevation").AsDouble(), 6);
                        double s2 = Math.Round(elem.LookupParameter("End Middle Elevation").AsDouble(), 6);
                        if (c.Length >= Ngd && s1 == s2)
                        {
                            ducts.Add(elem);
                        }
                    }
                }
                else if (elem.Category.Name == "Pipes")
                {
                    wsid = elem.WorksetId;
                    wsName = worksetnames[worksetIDs.IndexOf(wsid)];
                    if (GetSystemRank(wsName) == -1 || GetSystemRank(wsName) == 0) { continue; }
                    else
                    {
                        c = ((LocationCurve)elem.Location).Curve;
                        if (c.Length >= Ngd && Math.Abs(Math.Round(((Line)c).Direction.Normalize().Z, 3)) != 1)
                        {
                            pipes.Add(elem);
                        }
                    }
                }
                else if (elem is CableTray)
                {
                    c = ((LocationCurve)elem.Location).Curve;
                    if (c.Length >= Ngd && Math.Abs(Math.Round(((Line)c).Direction.Normalize().Z, 3)) == 0)
                    {
                        cables.Add(elem);
                    }
                }
            }
            #endregion

            if (ducts.Count == 0 && pipes.Count == 0 && cables.Count == 0)
            {
                TaskDialog.Show("Wrong Selection", "There are no Ducts or Pipes or Cabletrays selected.");
                return Result.Failed;
            }

            if (floors.Count == 0)
            {
                TaskDialog.Show("Error", "Linked Revit has not Floors to Host Hangers." + Environment.NewLine + "Make sure that Linked Revit is Structural discipline and has Floors.");
                return Result.Failed;
            }


            ductFitingFamInsFEC = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance));
            elemCatFilterDuctFit = new ElementCategoryFilter(BuiltInCategory.OST_DuctFitting);
            // ductFits was allDuctFits before shortening.
            ductFits = ductFitingFamInsFEC.WherePasses(elemCatFilterDuctFit)
                .Where(x => ((x.Location as LocationPoint).Point.Z) <= floorElevations.Max() && ((x.Location as LocationPoint).Point.Z) >= floorElevations.Min()).ToList();



            waterSuplyDiams = ListAdd(1)[0];
            waterSuplySpacers = ListAdd(1)[1];
            chillWaterDiams = ListAdd(2)[0];
            chillWaterSpacers = ListAdd(2)[1];
            drainageDiams = ListAdd(3)[0];
            drainageSpacers = ListAdd(3)[1];
            fireDiams = ListAdd(4)[0];
            fireSpacers = ListAdd(4)[1];

            List<List<double>> allDiameters = new List<List<double>>() { waterSuplyDiams, chillWaterDiams, drainageDiams, fireDiams };

            //===============================================================================================

            List<XYZ> HOs = new List<XYZ>();
            List<List<XYZ>> hangptsss = new List<List<XYZ>>();
            List<double> Belevs = new List<double>();
            List<double> Dhs = new List<double>();
            List<double> Dws = new List<double>();
            List<int> ffls = new List<int>();
            List<double> insthicks = new List<double>();
            IList<Element> DUCTS = new List<Element>();

            #region ducts
            foreach (Element d in ducts)
            {
                double s = 0;
                Curve c = ((LocationCurve)d.Location).Curve;
                List<XYZ> pts = new List<XYZ>();
                double Dw = 0;
                try { Dw = d.LookupParameter("Width").AsDouble(); }
                catch { Dw = d.LookupParameter("Diameter").AsDouble(); }
                int Dcoun = 0;
                if (AllWorksetsDIMS[0].Count == 2)
                {
                    s = Convert.ToDouble(AllWorksetsDIMS[0][1][0]);
                }
                else
                {
                    foreach (string qq in AllWorksetsDIMS[0][0])
                    {
                        double from = Convert.ToDouble(qq);
                        double to = Convert.ToDouble(AllWorksetsDIMS[0][1][Dcoun]);
                        double sp = Convert.ToDouble(AllWorksetsDIMS[0][2][Dcoun]);
                        if (Dw * 304.8 > from && Dw * 304.8 <= to)
                        {
                            s = sp / 304.8;
                            break;
                        }
                        Dcoun++;
                    }
                }
                if (s == 0)
                {
                    continue;
                }
                ElementId lvlid = d.LookupParameter("Reference Level").AsElementId();
                double Dh = 0;
                try { Dh = d.LookupParameter("Height").AsDouble(); }
                catch { Dh = d.LookupParameter("Diameter").AsDouble(); }
                double Belev = d.LookupParameter("Bottom Elevation").AsDouble() + ((Level)doc.GetElement(lvlid)).Elevation;
                double insothic = d.LookupParameter("Insulation Thickness").AsDouble();
                XYZ FO = ((Line)c).Direction.Normalize(); XYZ HO = new XYZ(-FO.Y, FO.X, FO.Z);
                XYZ dp = c.Evaluate(0.5, true);
                foreach (Element dft in ductFits)
                {
                    td("Duct fits considered");
                    Parameter p = dft.LookupParameter("Duct Width 1");
                    Parameter p2 = dft.LookupParameter("Takeoff Fixed Length"); double I = 0;
                    Parameter p3 = dft.LookupParameter("Angle"); double an = 0;
                    FamilyInstance fsl = dft as FamilyInstance;
                    if (p == null)
                    {
                        continue;
                    }
                    try
                    {
                        I = p2.AsDouble();
                        an = p3.AsDouble();
                    }
                    catch { }
                    double df = Math.Sqrt(Math.Pow(Dw, 2) + Math.Pow(p.AsDouble(), 2)) * 0.50;
                    XYZ dpi = ((LocationPoint)dft.Location).Point;
                    double diss = 0.50 * (Math.Tan(an) * I);
                    XYZ dpf = new XYZ(dpi.X, dpi.Y, dp.Z).Add(-diss * fsl.FacingOrientation);
                    Curve circ = Ellipse.CreateCurve(dpf, df, df, FO, HO, 0, 2 * Math.PI * df);
                    SetComparisonResult scr = circ.Intersect(c, out IntersectionResultArray ira);
                    if (ira != null)
                    {
                        if (ira.Size == 2)
                        {
                            XYZ Pdf0 = ira.get_Item(0).XYZPoint; XYZ Pdf1 = ira.get_Item(1).XYZPoint;
                            pts.Add(Pdf0); pts.Add(Pdf1);
                        }
                    }
                }

                XYZ P0 = c.Evaluate(0, true); XYZ Pf = c.Evaluate(1, true); List<XYZ> p0p1 = new List<XYZ>() { P0, Pf };
                List<XYZ> dps = DecOrder(pts, c); List<XYZ> pps = DecOrder(p0p1, c);
                XYZ dir = Line.CreateBound(pps[0], pps[1]).Direction.Normalize();
                XYZ Ps = pps[0].Add(f * dir); XYZ Pe = pps[1].Add(-f * dir); Curve cc = null;
                try
                {
                    cc = Line.CreateBound(Ps, Pe) as Curve;
                }
                catch { continue; }

                HOs.Add(HO); Dws.Add(Dw); Dhs.Add(Dh); insthicks.Add(insothic); Belevs.Add(Belev);
                List<XYZ> hangpts = new List<XYZ>();

                if (dps.Count == 0) // No ductfittings
                {
                    if (c.Length > Ngd && c.Length <= 4 * f)
                    {
                        hangpts.Add(dp);
                    }
                    else if (c.Length <= s && c.Length > 4 * f)
                    {
                        hangpts.Add(Ps);
                        hangpts.Add(Pe);
                    }
                    else if (c.Length > s)
                    {
                        hangpts.Add(Ps);
                        double n = Math.Ceiling(cc.Length / s) - 1;
                        XYZ Pis = Ps;
                        double Ns = (cc.Length / (n + 1));
                        for (int i = 0; i < n; i++)
                        {
                            XYZ Pie = Pis.Add(s * dir);
                            hangpts.Add(Pie);
                            Pis = Pie;
                        }
                        hangpts.Add(Pe);
                    }
                }
                else       // With ductfittings
                {
                    if (c.Length > Ngd && c.Length <= 4 * f)
                    {
                        List<XYZ> ps2 = new List<XYZ>() { dps[0], dps[1], dp };
                        List<XYZ> ps2o = DecOrder(ps2, c);
                        if (ps2o.IndexOf(dp) == 1)
                        {
                            hangpts.Add(ps2o[2]);
                        }
                        else
                        {
                            hangpts.Add(dp);
                        }
                    }
                    else if (c.Length <= s && c.Length > 4 * f)
                    {
                        List<XYZ> ps1 = new List<XYZ>() { dps[0], dps[1], Ps };
                        List<XYZ> ps1o = DecOrder(ps1, c);
                        if (ps1o.IndexOf(Ps) == 1)
                        {
                            hangpts.Add(ps1o[2]);
                        }
                        else
                        {
                            hangpts.Add(Ps);
                        }
                        List<XYZ> ptss = new List<XYZ>() { Ps, Pe };
                        for (int jj = 0; jj < 2; jj++)
                        {
                            for (int io = 0; io < dps.Count; io += 2)
                            {
                                List<XYZ> ps = new List<XYZ>() { dps[io], dps[io + 1], ptss[jj] };
                                List<XYZ> pso = DecOrder(ps, c);
                                if (pso.IndexOf(ptss[jj]) == 1) //Between Case
                                {
                                    if (jj == 0)
                                    {
                                        hangpts.Add(pso[2]);
                                    }
                                    else
                                    {
                                        hangpts.Add(pso[0]);
                                    }
                                    break;
                                }
                            }
                        }
                        hangpts.Add(Pe);
                    }
                    else if (c.Length > s)
                    {
                        List<XYZ> ps1 = new List<XYZ>() { dps[0], dps[1], Ps };
                        List<XYZ> ps1o = DecOrder(ps1, c);
                        if (ps1o.IndexOf(Ps) == 1)
                        {
                            hangpts.Add(ps1o[2]);
                        }
                        else
                        {
                            hangpts.Add(Ps);
                        }
                        double n = Math.Ceiling(cc.Length / s) - 1;
                        XYZ Pis = Ps;
                        for (int i = 0; i < n; i++)
                        {
                            XYZ Pie = Pis.Add(s * dir);
                            for (int io = 0; io < dps.Count; io += 2)
                            {
                                List<XYZ> ps3 = new List<XYZ>() { dps[io], dps[io + 1], Pie };
                                List<XYZ> ps3o = DecOrder(ps3, c);
                                if (ps3o.IndexOf(Pie) == 1) //Between Case
                                {
                                    Pie = ps3o[0];
                                    break;
                                }
                            }
                            hangpts.Add(Pie);
                            Pis = Pie;
                        }
                    }
                }
                hangptsss.Add(hangpts); DUCTS.Add(d);
            }
            #endregion

            List<double> HangDias1 = new List<double>() { 17, 22, 27, 34, 42, 52, 65, 67, 77, 82, 92, 102, 112, 127, 152, 162, 202 };
            List<double> HangDias2 = new List<double>() { 227, 252, 317, 352, 402 };

            List<double> pelevs = new List<double>(); List<double> pdias = new List<double>();
            List<double> emes = new List<double>(); List<double> smes = new List<double>();
            List<double> slps = new List<double>(); List<XYZ> HOps = new List<XYZ>();
            List<XYZ> PipePSs = new List<XYZ>(); List<ElementId> plvlids = new List<ElementId>();
            List<bool> Firebool = new List<bool>(); List<XYZ> FOps = new List<XYZ>();
            List<List<XYZ>> pangptsss = new List<List<XYZ>>(); List<Element> PIPES = new List<Element>();
            foreach (Element p in pipes)
            {

                WorksetId wsid = p.WorksetId;
                string ws = worksetnames[worksetIDs.IndexOf(wsid)];
                int Rank = GetSystemRank(ws);
                double dd = p.LookupParameter("Diameter").AsDouble() * 304.8; //mm
                double newdia = dd;
                //foreach (double d in Alldiameters[Rank - 1])
                //{
                //    if (Math.Round(dd, 3) <= d)
                //    {
                //        newdia = d; break;
                //    }
                //}
                //if (newdia == 0) { newdia = Alldiameters[Rank - 1].Last(); }
                double s = SysSpacing(Rank, newdia);
                if (s == 0)
                {
                    continue;
                }
                Curve c = ((LocationCurve)p.Location).Curve; double Ng = 100 / 304.8; double ff = 500 / 304.80;
                XYZ FOp = ((Line)c).Direction.Normalize(); XYZ HOp = new XYZ(-FOp.Y, FOp.X, FOp.Z);
                XYZ P0 = c.Evaluate(0, true); XYZ Pf = c.Evaluate(1, true); XYZ PipePS = P0; PipePSs.Add(PipePS);
                XYZ Ps = P0.Add(f * FOp); XYZ Pe = Pf.Add(-f * FOp); ;
                Curve cc = null;
                try
                {
                    cc = Line.CreateBound(Ps, Pe) as Curve;
                }
                catch { continue; }
                if (Rank == 4)
                {
                    Firebool.Add(true);
                }
                else
                {
                    Firebool.Add(false);
                }
                List<XYZ> pps = new List<XYZ>(); List<XYZ> pangpts = new List<XYZ>();
                double pdia = (newdia / 304.8) + (2 * p.LookupParameter("Insulation Thickness").AsDouble()); double HD = -1;
                foreach (double hd in HangDias1)
                {
                    if (pdia <= (hd / 304.8))
                    {
                        HD = hd / 304.8; break;
                    }
                }
                if (HD == -1)
                {
                    foreach (double hd in HangDias2)
                    {
                        if (pdia <= (hd / 304.8))
                        {
                            HD = hd / 304.8; break;
                        }
                    }
                }
                if (HD == -1) { HD = 402 / 304.8; }
                pdias.Add(HD);
                ElementId plvlid = p.LookupParameter("Reference Level").AsElementId(); plvlids.Add(plvlid);
                double OutSideD = p.LookupParameter("Outside Diameter").AsDouble();
                double pelev = c.GetEndPoint(0).Z - p.LookupParameter("Insulation Thickness").AsDouble(); pelevs.Add(pelev);
                double sme = p.LookupParameter("Start Middle Elevation").AsDouble(); smes.Add(sme);
                double eme = p.LookupParameter("End Middle Elevation").AsDouble(); emes.Add(eme);
                double slp = p.LookupParameter("Slope").AsDouble(); slps.Add(slp);

                if (c.Length > Ng && c.Length <= ff)
                {
                    CheckAdd(c.Evaluate(0.50, true), pangpts);
                }
                else if (c.Length <= s && c.Length > ff)
                {
                    CheckAdd(Ps, pangpts);
                    CheckAdd(Pe, pangpts);
                }
                else if (c.Length > s)
                {
                    CheckAdd(Ps, pangpts);
                    double n = Math.Ceiling(cc.Length / s) - 1;
                    XYZ Pis = Ps;
                    double Ns = (cc.Length / (n + 1));
                    for (int i = 0; i < n; i++)
                    {
                        XYZ Pie = Pis.Add(Ns * FOp);
                        CheckAdd(Pie, pangpts);
                        Pis = Pie;
                    }
                    CheckAdd(Pe, pangpts);
                }
                pangptsss.Add(pangpts); PIPES.Add(p);
                HOps.Add(FOp); FOps.Add(HOp);
            }

            List<double> widthes = new List<double>(); List<ElementId> plvids = new List<ElementId>();
            List<double> pelevsct = new List<double>(); List<List<XYZ>> CTPS = new List<List<XYZ>>();
            List<Element> CTS = new List<Element>();
            List<XYZ> HOcts = new List<XYZ>(); List<XYZ> FOcts = new List<XYZ>();
            foreach (Element p in cables)
            {
                Curve c = ((LocationCurve)p.Location).Curve; double Ng = 100 / 304.8; double ff = 500 / 304.80;
                XYZ FOp = ((Line)c).Direction.Normalize(); XYZ HOp = new XYZ(-FOp.Y, FOp.X, FOp.Z);
                XYZ P0 = c.Evaluate(0, true); XYZ Pf = c.Evaluate(1, true);
                XYZ Ps = P0.Add(f * FOp); XYZ Pe = Pf.Add(-f * FOp);
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
                    CheckAdd(c.Evaluate(0.50, true), pangpts);
                }
                else if (c.Length <= s && c.Length > ff)
                {
                    CheckAdd(Ps, pangpts);
                    CheckAdd(Pe, pangpts);
                }
                else if (c.Length > s)
                {
                    CheckAdd(Ps, pangpts);
                    double n = Math.Floor((cc.Length + (100 / 304.8)) / s);
                    XYZ Pis = Ps;
                    for (int i = 0; i < n; i++)
                    {
                        XYZ Pie = Pis.Add(s * FOp);
                        CheckAdd(Pie, pangpts);
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
                foreach (Element d in DUCTS)
                {
                    if (hangptsss[h].Count > 0)
                    {
                        foreach (XYZ p in hangptsss[h])
                        {
                            int fflu = -1; int ffld = -1; double ROD = 0;
                            double dd = double.MinValue; double dd2 = double.MaxValue; double Zu = 0; double Zd = 0;
                            foreach (Face face in floorfacesd)
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
                                            fflu = floorfacesd.IndexOf(face); Zu = ip.Z - p.Z;
                                        }
                                        else if (D > 0 && D <= dd2)
                                        {
                                            if (D < dd2)
                                            {
                                                dd2 = D;
                                                ffld = floorfacesd.IndexOf(face); Zd = p.Z - ip.Z;
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
                            XYZ PP = new XYZ(p.X, p.Y, p.Z - insthicks[h] - (Dhs[h] / 2));
                            FamilyInstance hang = doc.Create.NewFamilyInstance(PP, ductHanger, HOs[h], d, StructuralType.NonStructural);
                            hang.LookupParameter("Width").Set(Dws[h] + (2 * insthicks[h]) + 16 / 304.8);
                            //if (FS == fsd1)
                            //{
                            double Z = Belevs[h] - insthicks[h] - hang.LookupParameter("Elevation from Level").AsDouble();
                            hang.Location.Move(new XYZ(0, 0, Z));
                            ROD += insthicks[h] + (Dhs[h] / 2);
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


                int pir = 0;
                foreach (Element pip in PIPES)
                {
                    int hangco = 0;
                    foreach (XYZ p in pangptsss[pir])
                    {
                        int fflu = -1; int ffld = -1; double ddu = double.MinValue; double ddd = double.MaxValue;
                        double Zd = 0; double Zu = 0; XYZ IP = null;
                        foreach (Face face in floorfacesd)
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
                                    if (D > 0 && D <= ddd)
                                    {
                                        ddd = D;
                                        ffld = floorfacesd.IndexOf(face); Zd = ip.Z; IP = ip;
                                    }
                                    else if (D < 0 && D >= ddu)
                                    {
                                        ddu = D;
                                        fflu = floorfacesd.IndexOf(face); Zu = ip.Z; IP = ip;
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
                        double hu = Math.Abs(Zu - p.Z); double hd = Math.Abs(p.Z - Zd);
                        Face facee = null; int ffl = 0;
                        if (hu < hd && fflu != -1)
                        {
                            if (Firebool[pir])
                            {
                                FamilyInstance pangq = doc.Create.NewFamilyInstance(p, pipeHanger2, FOps[pir], doc.GetElement(plvlids[pir]), StructuralType.NonStructural); double q = 1;
                                if (emes[pir] < smes[pir])
                                {
                                    q = -1;
                                }
                                double ppeleve = pelevs[pir] + (q * slps[pir] * p.DistanceTo(PipePSs[pir]));
                                double pofl = ppeleve - (3000 / 304.8);
                                pangq.LookupParameter("Diameter").Set(pdias[pir]);
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
                                facee = floorfacesd[fflu]; ffl = fflu;
                            }
                        }
                        else if (hu > hd && ffld != -1)
                        {
                            facee = floorfaces[ffld]; ffl = ffld;
                        }
                        if (!Firebool[pir] && facee != null)
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
                                if (pdias[pir] > (202 / 304.8))
                                {
                                    FS = pipeHanger200;
                                }
                                else { FS = pipeHanger20; }
                                FS.Activate();
                                FamilyInstance pang = doc.Create.NewFamilyInstance(reffface, p, HOps[pir], FS);
                                pang.LookupParameter("Schedule Level").Set(floors[ffl].LevelId);
                                Line ll = Line.CreateUnbound(p, XYZ.BasisZ);
                                double rr = pang.HandOrientation.AngleOnPlaneTo(HOps[pir], XYZ.BasisZ);
                                IntersectionResultArray iraa = new IntersectionResultArray();
                                SetComparisonResult scr = facee.Intersect(ll, out iraa);
                                if (iraa != null && !iraa.IsEmpty)
                                {
                                    Curve cv = Line.CreateBound(p, iraa.get_Item(0).XYZPoint);
                                    pang.LookupParameter("Pipe_distance").Set(cv.Length - (0.5 * pdias[pir]));
                                }
                                pang.LookupParameter("Pipe Outer Diameter").Set(pdias[pir]);
                            }
                            catch
                            {
                                if (hangco == 0)
                                {
                                    errco++;
                                    err += pip.Id + "\n";
                                }
                            }
                        }
                        hangco++;
                    }
                    pir++;
                }
                int ctco = 0;
                foreach (Element ct in CTS)
                {
                    foreach (XYZ p in CTPS[ctco])
                    {
                        int fflu = -1; double ddu = double.MinValue;
                        double Zu = 0;
                        foreach (Face face in floorfacesd)
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
                                        fflu = floorfacesd.IndexOf(face); Zu = ip.Z;
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
        List<XYZ> DecOrder(List<XYZ> oldlist, Curve cu)
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

        double SysSpacing(int rankk, double dia)
        {
            double spac = 0;
            if (rankk == 1)
            {
                spac = GetSpace(waterSuplyDiams, waterSuplySpacers, dia);
            }
            else if (rankk == 2)
            {
                spac = GetSpace(chillWaterDiams, chillWaterSpacers, dia);
            }
            else if (rankk == 3)
            {
                spac = GetSpace(drainageDiams, drainageSpacers, dia);
            }
            else if (rankk == 4)
            {
                spac = GetSpace(fireDiams, fireSpacers, dia);
            }
            return spac;
        }

        Tuple<bool, Solid, double> processFloor(Element floor)
        {
            Options options = app.NewGeometryOptions();
            options.ComputeReferences = true;
            GeometryElement geoElem = floor.get_Geometry(options);
            foreach (GeometryObject geo in geoElem)
            {
                Solid solid = geo as Solid;
                if (solid != null && solid.Volume != 0)
                {
                    double elevAtBot = 0, elevAtTop = 0;
                    Parameter atBotParam = floor.LookupParameter("Elevation at Bottom");
                    Parameter atTopParam = floor.LookupParameter("Elevation at Top");
                    if (atBotParam != null && atTopParam != null)
                    {
                        elevAtBot = atBotParam.AsDouble(); elevAtTop = atTopParam.AsDouble();
                        if (elevAtBot == 0 && elevAtTop == 0)
                        {
                            double levelElevation = 0;
                            Parameter levelParam = floor.LookupParameter("Level");
                            if (levelParam != null)
                            {
                                levelElevation = lvls.Where(l => l.Name.Trim() == levelParam.AsValueString().Trim()).FirstOrDefault().Elevation;
                                double offset = floor.LookupParameter("Height Offset From Level").AsDouble();
                                elevAtBot = levelElevation + offset;
                            }
                        }
                        return new Tuple<bool, Solid, double>(true, solid, elevAtBot);
                    }
                }
            }
            return new Tuple<bool, Solid, double>(false, null, 0);
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
        Face GetFaces(Element ele)
        {
            Face face = null;
            Options optt = app.NewGeometryOptions();
            optt.View = doc.ActiveView;
            Solid s1 = null;
            optt.ComputeReferences = true;
            GeometryElement gele = ele.get_Geometry(optt);
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
        void CheckAdd(XYZ p, List<XYZ> pss)
        {
            if (!pss.Contains(p))
            {
                pss.Add(p);
            }
        }
        DialogResult td(object g)
        {
            return MessageBox.Show(g + " ");
        }

    }
    public class selcFilterPDCt : ISelectionFilter
    {
        public bool AllowElement(Element e)
        {
            if (e.Category != null)
            {
                if (e is Pipe || e is Duct || e is CableTray)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            throw new NotImplementedException();
        }
    }
}

