using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IBIMSGen.Penetration
{
    [Transaction(TransactionMode.Manual)]
    public class Penetration : IExternalCommand
    {
        UIDocument uidoc;
        Document doc;
        Autodesk.Revit.Creation.Application app;
        IList<Level> levels;
        IList<Element> AllLinks, allLinksUnique, aLLWalls, ALLBEAMS, selectedLinks, ductsCablesConduitsPipes, floordrains, ducts, cables, conduits, pipes, floors;
        IList<HostElement> hostElements;
        List<string> familysymbolnamesuniq, selDocsNames, linkchars;
        List<double> diasDR;
        FamilySymbol dr, ws, chw, ff, con, fssd, fssct, fssDR;
        IList<Reference> references;
        StringBuilder sb;
        Options optt;
        int FSCount;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet element)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = commandData.Application.ActiveUIDocument.Document;
            app = commandData.Application.Application.Create;
            diasDR = new List<double>() { 15, 20, 25, 32, 40, 50, 75, 110, 160, 200, 250, 315, 355 };
            sb = new StringBuilder();
            optt = app.NewGeometryOptions();
            optt.ComputeReferences = true;
            FilteredElementCollector allsymbols = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_GenericModel);
            FSCount = 0;
            #region family names

            List<string> floordrainfamilies = new List<string>() { "01- Area Drain", "01- FLOOR DRAIN", "1- FFD" , "02- SHOWER DRAIN", "Intercepting_Drain-Watts-FD-440_Series", "KA_FD",
            "SABPLMG_1160_PD_PFIX_Area Drain 1","SABPLMG_1160_PD_PFIX_Area Drain 2","SABPLMG_1160_PD_PFIX_Funnel Floor Drain Type2","SABPLMG_1160_PD_PFIX_Parking Drain","SABPLMG_1160_PD_PFIX_Shower Floor Drain"};

            #endregion

            dr = null; ws = null; chw = null; ff = null; con = null; fssd = null; fssct = null; fssDR = null;

            FilteredWorksetCollector worksetFilteredCollector = new FilteredWorksetCollector(doc);
            List<string> worksetS = new List<string>();
            List<Workset> worksets = new List<Workset>();
            List<WorksetId> worksetIDs = new List<WorksetId>();

            #region get worksets
            worksetIDs.Add(worksetFilteredCollector.Where(x => x.Name.ToLower().Contains("drainage")).Select(x => x.Id).FirstOrDefault());
            worksetIDs.Add(worksetFilteredCollector.Where(x => x.Name.ToLower().Contains("water")).Select(x => x.Id).FirstOrDefault());
            worksetIDs.Add(worksetFilteredCollector.Where(x => x.Name.ToLower().Contains("hvac")).Select(x => x.Id).FirstOrDefault());
            worksetIDs.Add(worksetFilteredCollector.Where(x => x.Name.ToLower().Contains("fire")).Select(x => x.Id).FirstOrDefault());
            worksetIDs.Add(worksetFilteredCollector.Where(x => x.Name.ToLower().Contains("duct")).Select(x => x.Id).FirstOrDefault());
            worksetIDs.Add(worksetFilteredCollector.Where(x => x.Name.ToLower().Contains("ltpw")).Select(x => x.Id).FirstOrDefault());
            #endregion

            #region get family symbols
            foreach (FamilySymbol fs in allsymbols)
            {
                if (fs.Family.Name == "MSAR_MHT_ST_CIRC_OPENING")
                {
                    if (fs.Name == "DR")
                    {
                        dr = fs;
                    }
                    else if (fs.Name == "WS")
                    {
                        ws = fs;
                    }
                    else if (fs.Name == "CHW")
                    {
                        chw = fs;
                    }
                    else if (fs.Name == "FP")
                    {
                        ff = fs;
                    }
                    else if (fs.Name == "COND")
                    {
                        con = fs;
                    }
                }
                else if (fs.Family.Name == "MSAR_MHT_ST_REC_OPENING")
                {
                    if (fs.Name == "Duct")
                    {
                        fssd = fs;
                    }
                    else if (fs.Name == "CT")
                    {
                        fssct = fs;
                    }
                    else if (fs.Name == "DR")
                    {
                        fssDR = fs;
                    }
                }
            }

            List<FamilySymbol> FSS = new List<FamilySymbol>() { dr, ws, chw, ff, con };

            foreach (FamilySymbol fy in FSS)
            {
                if (fy == null || fssd == null)
                {
                    TaskDialog.Show("Load Families", "Please Load Families First");
                    return Result.Failed;
                }
            }
            #endregion
            familysymbolnamesuniq = allsymbols.Cast<FamilySymbol>().Select(fi => fi.Family.Name).ToList();

            levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().OrderBy(x => x.Elevation).ToList();
            AllLinks = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).ToList();

            allLinksUnique = new List<Element>();

            if (AllLinks.Count == 0)
            {
                td("No Links Loaded!!");
                return Result.Failed;
            }
            PenetrationForm form = new PenetrationForm(AllLinks.Select(x => doc.GetElement(x.GetTypeId()).Name).Distinct().ToList(),
                allsymbols.Cast<FamilySymbol>().ToList(), familysymbolnamesuniq);

            allLinksUnique = new List<Element>();
            foreach (Element link in AllLinks)
            {
                if (!allLinksUnique.Contains(link))
                {
                    allLinksUnique.Add(link);
                }
            }
            td(allLinksUnique.Count.ToString());
            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel) { return Result.Cancelled; }
            if (form.radioButton1.Checked)
            {
                WorksetsForm wsform = new WorksetsForm(worksetFilteredCollector.Select(x=>x.Name).ToList(),allsymbols.Cast<FamilySymbol>().ToList());
                wsform.Show();
                return Result.Succeeded;
            }
            else
            {

                selectedLinks = new List<Element>();
                selDocsNames = new List<string>();
                td("test");
                foreach (int i in form.linksinds)
                {
                    string linkPath = null;
                    linkPath = ((RevitLinkInstance)allLinksUnique[i]).GetLinkDocument()?.PathName?.Trim();
                    if (linkPath == null) continue;
                    selectedLinks.Add(allLinksUnique[i]);
                    selDocsNames.Add(linkPath);
                }
                hostElements = new List<HostElement>();
                floors = new List<Element>();
                ductsCablesConduitsPipes = new List<Element>();
                floordrains = new List<Element>();
                linkchars = new List<string>();
                if (form.sel)
                {
                    references = uidoc.Selection.PickObjects(ObjectType.LinkedElement, new selectionFilter(), "Select Elements in RVT Link");
                    foreach (Reference reference in references)
                    {
                        Document linkedDoc = ((RevitLinkInstance)doc.GetElement(reference.ElementId)).GetLinkDocument();
                        if (!selDocsNames.Contains(linkedDoc.PathName.Trim())) continue;
                        Element linkedElement = linkedDoc.GetElement(reference.LinkedElementId);
                        string linkedTypeName = doc.GetElement(((RevitLinkInstance)doc.GetElement(reference.ElementId)).GetTypeId()).Name;
                        if (linkedElement is Pipe || linkedElement is Duct || linkedElement is CableTray || linkedElement is Conduit)
                        {
                            ductsCablesConduitsPipes.Add(linkedElement);
                            linkchars.Add(linkedTypeName.Split('_')[6]);
                        }
                        else if (linkedTypeName.Split('_').Count() >= 7 && linkedElement is FamilyInstance)
                        {
                            if (linkedTypeName.Split('_')[6] == "DRGW" && floordrainfamilies.Contains(((FamilyInstance)linkedElement).Symbol.Family.Name))
                            {
                                floordrains.Add(linkedElement);
                            }
                        }
                        else if (linkedElement is Floor || linkedElement is Wall)
                        {
                            RevitLinkInstance rli = (RevitLinkInstance)doc.GetElement(reference.ElementId);
                            hostElements.Add(new HostElement(linkedElement, rli));
                        }
                    }
                }
                else
                {
                    getLinkedElements();

                }

                IList<PenetratingElement> penetratingElements = new List<PenetratingElement>();
                for (int i = 0; i < ductsCablesConduitsPipes.Count; i++)
                {
                    Element elem = ductsCablesConduitsPipes.ElementAt(i); //element
                    LocationCurve elementLocationCurve = ((LocationCurve)elem.Location);
                    Curve elementCurve = null;
                    if (elementLocationCurve != null) { elementCurve = elementLocationCurve.Curve; }
                    else { continue; }
                    double elementWidth = 0, elementHeight = 0, insulationThickness = 0;
                    double diameter = 0;
                    WorksetId worksetId = null;
                    FamilySymbol familySymbol = null; // dr ws chw ff con 

                    if (elem is Pipe || elem is Conduit)
                    {
                        double outsideDiam = 0;
                        if (elem is Pipe)
                        {
                            string name = linkchars[i];
                            if (name == "DRGW")
                            {
                                familySymbol = dr;
                                worksetId = worksetIDs[0];
                            }
                            else if (name == "WATR")
                            {
                                familySymbol = ws;
                                worksetId = worksetIDs[1];
                            }
                            else if (name == "PIPE")
                            {
                                familySymbol = chw;
                                worksetId = worksetIDs[2];
                            }
                            else if (name == "FIRE")
                            {
                                familySymbol = ff;
                                worksetId = worksetIDs[3];
                            }
                            else
                            {
                                familySymbol = dr;
                                worksetId = worksetIDs[0];
                            } /// Not Defined
                            outsideDiam = elem.LookupParameter("Outside Diameter").AsDouble() + 2 * (elem.LookupParameter("Insulation Thickness").AsDouble());
                        }
                        else
                        {
                            outsideDiam = elem.LookupParameter("Outside Diameter").AsDouble();
                            familySymbol = FSS[4];
                            worksetId = worksetIDs[5];
                        }

                        // loop in steel pipes 
                        for (int k = 0; k < diasDR.Count; k++)
                        {
                            double d = diasDR[k];
                            if (Math.Round(outsideDiam * 304.8) <= d)
                            {
                                if (Math.Round(outsideDiam * 304.8) == d)
                                {
                                    if (k != diasDR.Count - 1)
                                    {
                                        diameter = diasDR[k + 1] / 304.8;
                                    }
                                    else
                                    {
                                        diameter = d / 304.8;
                                    }
                                }
                                else
                                {
                                    diameter = d / 304.8;
                                }
                                break;
                            }
                        }
                        penetratingElements.Add(new PenetratingElement(elem, worksetId, diameter, diameter, 0, elementCurve, familySymbol));
                    }
                    else
                    {
                        try
                        {
                            elementWidth = elem.LookupParameter("Width").AsDouble();
                            elementHeight = elem.LookupParameter("Height").AsDouble();
                        }
                        catch { continue; }
                        try
                        {
                            insulationThickness = elem.LookupParameter("Insulation Thickness").AsDouble();
                        }
                        catch { }
                        if (elem is CableTray)
                        {
                            familySymbol = fssct;
                            worksetId = worksetIDs[5];
                        }
                        else
                        {
                            familySymbol = fssd;
                            worksetId = worksetIDs[4];
                        }
                        penetratingElements.Add(new PenetratingElement(elem, worksetId, elementWidth, elementHeight, insulationThickness, elementCurve, familySymbol));
                    }
                    if (familySymbol == null)
                    {
                        familySymbol = FSS[0];//familysymbol
                    }
                }
                foreach (Element floorDrain in floordrains)
                {
                    XYZ locationPoint = ((LocationPoint)floorDrain.Location).Point;
                    Curve curve = Line.CreateBound(locationPoint.Add(XYZ.BasisZ * 4), locationPoint.Add(-XYZ.BasisZ * 4));
                    penetratingElements.Add(new PenetratingElement(floorDrain, worksetIDs[0], 200 / 304.8, 200 / 304.8, 0, curve, fssDR));
                }




                ProgressBar progressBarForm = new ProgressBar();
                progressBarForm.progresBarRatio.Minimum = 0;
                progressBarForm.progresBarRatio.Maximum = penetratingElements.Count;
                progressBarForm.progresBarRatio.Step = 1;
                progressBarForm.Lb.Text = "0 / " + penetratingElements.Count;
                progressBarForm.Show();

                List<XYZ> locs = new List<XYZ>();
                td(hostElements.Count.ToString());

                #region transaction group
                TransactionGroup tg = new TransactionGroup(doc);
                tg.SetName("Penetration");
                tg.Start();

                #region transaction
                using (Transaction tr = new Transaction(doc, "Pentration"))
                {
                    tr.Start();
                    #region activate family symbol
                    foreach (FamilySymbol fy in FSS)
                    {
                        fy.Activate();
                    }
                    fssd.Activate();
                    fssct.Activate();
                    #endregion

                    foreach (PenetratingElement penetratingElement in penetratingElements)
                    {
                        foreach (HostElement hostElement in hostElements)
                        {
                            Element elem = hostElement.element;
                            #region get solid

                            Solid so = getSolid(elem);

                            #endregion

                            #region progress bar update
                            if (so == null)
                            {
                                progressBarForm.Lb.Text = $"Progress: {hostElements.IndexOf(hostElement)}  / {hostElements.Count}(Task {penetratingElements.IndexOf(penetratingElement)} of {penetratingElements.Count})";
                                progressBarForm.progresBarRatio.PerformStep();
                                progressBarForm.progresBarRatio.Refresh();
                                progressBarForm.Refresh();
                                continue;
                            }
                            #endregion

                            XYZ sleeveDir = null;
                            Face face = null;
                            double depth = 0;
                            XYZ locationPt = null;
                            XYZ dir = null;
                            SolidCurveIntersectionOptions scio = new SolidCurveIntersectionOptions();
                            scio.ResultType = SolidCurveIntersectionMode.CurveSegmentsInside;
                            SolidCurveIntersection solidIntersectCurve = null;
                            try
                            {
                                solidIntersectCurve = so.IntersectWithCurve(penetratingElement.axis, scio);
                            }
                            catch { }
                            if (solidIntersectCurve == null) { continue; }
                            else
                            {
                                if (solidIntersectCurve.SegmentCount > 0)
                                {
                                    #region get depth of penetration
                                    Curve cc = solidIntersectCurve.GetCurveSegment(0);
                                    if (elem is Wall)
                                    {
                                        depth = ((Wall)elem).Width; //depth
                                    }
                                    else if (elem is Floor)
                                    {
                                        depth = elem.LookupParameter("Thickness").AsDouble();//depth
                                    }
                                    else
                                    {
                                        depth = cc.Length;//depth
                                    }
                                    if (cc is Line)
                                    {
                                        if (Math.Round(Math.Abs(((Line)cc).Direction.Normalize().Z), 1) == 1) //Floor
                                        {
                                            sleeveDir = XYZ.BasisX;//sleeveDir
                                        }
                                        else //Wall
                                        {
                                            sleeveDir = XYZ.BasisZ;//sleeveDir
                                        }
                                    }
                                    #endregion

                                    #region direction of penetration
                                    XYZ intersectionStartPt = solidIntersectCurve.GetCurveSegment(0).GetEndPoint(0);
                                    XYZ intersectionEndPt = solidIntersectCurve.GetCurveSegment(0).GetEndPoint(1);
                                    XYZ curveStartPt = penetratingElement.axis.GetEndPoint(0);
                                    XYZ curveEndPt = penetratingElement.axis.GetEndPoint(1);
                                    if (intersectionStartPt.X != curveStartPt.X && intersectionStartPt.Y != curveStartPt.Y
                                        && intersectionStartPt.X != curveEndPt.X && intersectionStartPt.Y != curveEndPt.Y)
                                    {
                                        locationPt = intersectionStartPt;//locationPoint
                                        dir = Line.CreateBound(intersectionStartPt, intersectionEndPt).Direction.Normalize();
                                    }
                                    else
                                    {
                                        locationPt = intersectionEndPt;//locationPoint
                                        dir = Line.CreateBound(intersectionEndPt, intersectionStartPt).Direction.Normalize();
                                    }
                                    #endregion

                                    #region get intersection face
                                    foreach (Face f in so.Faces)
                                    {
                                        IntersectionResultArray ira = new IntersectionResultArray();
                                        SetComparisonResult scr = f.Intersect(penetratingElement.axis, out ira);
                                        if (ira != null)
                                        {
                                            if (!ira.IsEmpty)
                                            {
                                                face = f;
                                                break;
                                            }
                                        }
                                    }
                                    #endregion

                                }
                            }

                            if (face != null && penetratingElement.familySymbol != null)
                            {
                                if (face.Reference != null)
                                {
                                    Reference faceRef = face.Reference.CreateLinkReference(hostElement.rli);//reference
                                    try
                                    {
                                        bool found = false;
                                        foreach (XYZ point in locs)
                                        {
                                            if (Math.Round(point.X, 2) == Math.Round(locationPt.X, 2)
                                                && Math.Round(point.Y, 2) == Math.Round(locationPt.Y, 2) & Math.Round(point.Z, 2) == Math.Round(locationPt.Z, 2))
                                            {
                                                found = true;
                                                break;
                                            }
                                        }
                                        if (!found)
                                        {
                                            FamilyInstance familyInstance = doc.Create.NewFamilyInstance(faceRef, locationPt, sleeveDir, penetratingElement.familySymbol);
                                            if (familyInstance != null)
                                            {
                                                if (penetratingElement.element is Pipe || penetratingElement.element is Conduit)
                                                {
                                                    familyInstance.LookupParameter("Diameter").Set(penetratingElement.width);
                                                }
                                                else
                                                {
                                                    familyInstance.LookupParameter("b").Set(penetratingElement.width + 2 * (penetratingElement.insulationThickness + (50 / 304.8)));
                                                    familyInstance.LookupParameter("h").Set(penetratingElement.height + 2 * (penetratingElement.insulationThickness + (50 / 304.8)));
                                                    familyInstance.Location.Rotate(((Line)penetratingElement.axis), Math.PI / 2);
                                                }
                                                FSCount++;
                                                locs.Add(locationPt);
                                                familyInstance.LookupParameter("Depth").Set(depth + (50 / 304.8));
                                                familyInstance.LookupParameter("Schedule Level").Set(ElevationSleeve(locationPt).Id);
                                                familyInstance.LookupParameter("Comments").Set(ElevationSleeve(locationPt).Name.ToString());
                                                familyInstance.LookupParameter("Workset").Set(penetratingElement.worksetId.IntegerValue);
                                            }
                                        }
                                    }
                                    catch { }
                                }
                            }
                            progressBarForm.Lb.Text = $"Progress: {hostElements.IndexOf(hostElement)}  / {hostElements.Count}(Task {penetratingElements.IndexOf(penetratingElement)} of {penetratingElements.Count})";
                            progressBarForm.progresBarRatio.PerformStep();
                            progressBarForm.progresBarRatio.Refresh();
                            progressBarForm.Refresh();
                        }
                    }
                    tr.Commit();
                }

                #endregion

                progressBarForm.Close();
                td("Finished !!" + "\n" + "All voids count is:  " + FSCount);

                tg.Assimilate();
                #endregion

                return Result.Succeeded;
            }
        }

        private Solid getSolid(Element elem)
        {
            GeometryElement gele = elem.get_Geometry(optt);
            foreach (GeometryObject geo in gele)
            {
                Solid g = geo as Solid;
                if (g != null && g.Volume != 0)
                {
                    return g;
                }
            }
            return null;
        }

        private void getLinkedElements()
        {
            foreach (Element linkInstance in selectedLinks)
            {
                RevitLinkInstance rli = linkInstance as RevitLinkInstance;
                Document linkedDoc = rli.GetLinkDocument();
                if (linkedDoc != null)
                {
                    floors = new FilteredElementCollector(linkedDoc).OfClass(typeof(Floor)).ToList();
                    aLLWalls = new FilteredElementCollector(linkedDoc).OfClass(typeof(Wall)).Where(x => ((Wall)x).Width >= (50 / 304.8)).ToList();
                    ALLBEAMS = new FilteredElementCollector(linkedDoc).OfClass(typeof(FamilyInstance)).Where(x => x.Category.Name == "Structural Framing").ToList();
                    ducts = new FilteredElementCollector(linkedDoc).OfClass(typeof(Duct)).ToList();
                    cables = new FilteredElementCollector(linkedDoc).OfClass(typeof(CableTray)).ToList();
                    conduits = new FilteredElementCollector(linkedDoc).OfClass(typeof(Conduit)).ToList();
                    pipes = new FilteredElementCollector(linkedDoc).OfClass(typeof(Pipe)).ToList();
                    foreach (Element e in ducts)
                    {
                        ductsCablesConduitsPipes.Add(e);
                        linkchars.Add(doc.GetElement(rli.GetTypeId()).Name.Split('_')[6]);
                    }
                    foreach (Element e in cables)
                    {
                        ductsCablesConduitsPipes.Add(e);
                        linkchars.Add(doc.GetElement(rli.GetTypeId()).Name.Split('_')[6]);
                    }
                    foreach (Element e in pipes)
                    {
                        ductsCablesConduitsPipes.Add(e);
                        linkchars.Add(doc.GetElement(rli.GetTypeId()).Name.Split('_')[6]);
                    }
                    foreach (Element e in conduits)
                    {
                        ductsCablesConduitsPipes.Add(e);
                        linkchars.Add(doc.GetElement(rli.GetTypeId()).Name.Split('_')[6]);
                    }
                    foreach (Element el in aLLWalls)
                    {
                        hostElements.Add(new HostElement(el, rli));
                    }
                    foreach (Element el in ALLBEAMS)
                    {
                        hostElements.Add(new HostElement(el, rli));
                    }
                    foreach (Element el in floors)
                    {
                        hostElements.Add(new HostElement(el, rli));
                    }
                }
            }
        }

        DialogResult td(object g)
        {
            return MessageBox.Show(g + " ");
        }

        Level ElevationSleeve(XYZ Z)
        {
            int ilvl = 0; Level lv = null;
            foreach (Level lvl in levels)
            {
                if (lvl.Elevation - Z.Z >= 0)
                {
                    if (ilvl == 0)
                    {
                        lv = levels[0];
                        break;
                    }
                    else
                    {
                        lv = levels[ilvl - 1];
                        break;
                    }
                }
                ilvl++;
            }
            if (lv == null) { lv = levels.Last(); }
            return lv;
        }

        bool isIntersected(IList<PenetratingElement> elements, Element penetratedElement)
        {
            Solid so = getSolid(penetratedElement);
            if (so == null) return false;
            SolidCurveIntersectionOptions scio = new SolidCurveIntersectionOptions();
            scio.ResultType = SolidCurveIntersectionMode.CurveSegmentsInside;
            SolidCurveIntersection solidIntersectCurve = null;
            foreach (PenetratingElement element in elements)
            {

                try
                {
                    solidIntersectCurve = so.IntersectWithCurve(element.axis, scio);
                }
                catch { return false; }
                if (solidIntersectCurve != null)
                {
                    if (solidIntersectCurve.SegmentCount > 0)
                    {
                        Curve cc = solidIntersectCurve.GetCurveSegment(0);
                        return true;
                    }
                }
            }
            return false;
        }
    }

}
