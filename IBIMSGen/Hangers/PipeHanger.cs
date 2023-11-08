using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace IBIMSGen.Hangers
{
    internal class PipeHanger : Hanger
    {
        public List<double> Diameters { get; }
        public RevitLinkInstance DocumentRLI { get; }

        public XYZ startPt { get; private set; }
        public bool isFireFighting { get; private set; }
        public ElementId levelId { get; private set; }
        public double OutsideDiam { get; private set; }
        public double pipeElevation { get; private set; }
        public double midElevStart { get; private set; }
        public double midElevEnd { get; private set; }
        public double slope { get; private set; }
        public XYZ pipeDirection { get; private set; }
        public Face LowerFace { get; private set; }
        public FamilySymbol FamilySymbol2 { get; private set; }
        public List<List<string>> AllLinksNames { get; private set; }
        public double HangerDiameter { get; private set; }
        public double PipeDiam { get; private set; }

        public PipeHanger(Document document, Solid solid, Element element, double up, double down,
            List<List<Dictionary<string, double>>> dimensions, List<FamilySymbol> symbols, double negligible,
            double offset, List<double> diameters, List<List<string>> linksNames, RevitLinkInstance linkInstance,
            QuadTree allDuctsTree, QuadTree allTraysTree, RevitLinkInstance pipesRLI = null)
        {
            Document = document;
            Solid = solid;
            Element = element;
            Up = up;
            Down = down;
            Dimensions = dimensions;
            Symbols = symbols;
            Negligible = negligible;
            Offset = offset;
            Diameters = diameters;
            AllLinksNames = linksNames;
            DocumentRLI = linkInstance;
            LinkInstance = pipesRLI;
            AllDuctsTree = allDuctsTree;
            AllTraysTree = allTraysTree;
            Process();
        }

        public void Process()
        {
            LowerFace = Solid.Faces.get_Item(0);
            FamilySymbol = null;
            FamilySymbol2 = null;
            Supports = new List<Support>();
            isFireFighting = false;
            try
            {
                Width = Element.LookupParameter("Diameter").AsDouble() * 304.8; //mm
            }
            catch
            {
                return;
            }
            ElementCurve = ((LocationCurve)Element.Location).Curve;
            GetRegion();
            //elementsNearby=
            pipeDirection = ((Line)ElementCurve).Direction.Normalize();
            Perpendicular = new XYZ(-pipeDirection.Y, pipeDirection.X, pipeDirection.Z);
            startPt = ElementCurve.Evaluate(0, true);
            XYZ Pf = ElementCurve.Evaluate(1, true);
            XYZ Ps = startPt.Add(Offset * pipeDirection);
            XYZ Pe = Pf.Add(-Offset * pipeDirection);
            Curve hangCurve = null;
            int Rank = -1;
            try
            {
                hangCurve = Line.CreateBound(Ps, Pe);
            }
            catch
            {
                return;
            }
            if (LinkInstance != null)
            {
                RevitLinkType rlt = Document.GetElement(LinkInstance.GetTypeId()) as RevitLinkType;
                try
                {

                    Rank = GetSystemRank(rlt.Name, true, AllLinksNames).Where(x => x != -1 && x != 0 && x != 5).First();
                }
                catch { return; }
                if (Rank == -1) return;
                Spacing = GetSysSpacing(Dimensions[Rank], Width);

                if (Spacing == 0) return;
                string familySymbolName = Symbols.Select(x => x.FamilyName).Distinct().ElementAt(Convert.ToInt32(Dimensions[Rank][0]["family"]));
                FamilySymbol = Symbols.Where(x => x.FamilyName.Equals(familySymbolName)).First();
                isFireFighting = Dimensions[Rank][0]["FF"] == 1;
                if (!isFireFighting)
                {
                    string familySymbolName2 = Symbols.Select(x => x.FamilyName).Distinct().ElementAt(Convert.ToInt32(Dimensions[Rank][0]["family2"]));
                    FamilySymbol2 = Symbols.Where(x => x.FamilyName.Equals(familySymbolName2)).First();
                }

            }
            else
            {
                int r = -1;
                try
                {
                    r = GetSystemRank("used", true, AllLinksNames).Where(x => x != -1 && x != 0 && x != 5).FirstOrDefault();
                }
                catch
                {
                    return;
                }
                if (r == -1 || r == 0 || r == 5) return;
                Spacing = GetSysSpacing(Dimensions[r], Width);
                if (Spacing == 0) return;

                isFireFighting = Dimensions[r].First()["FF"] == 1;
                int index = Convert.ToInt32(Dimensions[r].Where(x => x["spacing"] != 0).First()["family"]);
                if (index == -1) return;
                string familySymbolName = Symbols.Select(x => x.FamilyName).Distinct().ElementAt(index);
                FamilySymbol = Symbols.Where(x => x.FamilyName.Equals(familySymbolName)).First();
                int index2 = Convert.ToInt32(Dimensions[r].Where(x => x["spacing"] != 0).First()["family2"]);
                if (index2 >= 0)
                {
                    string familySymbolName2 = Symbols.Select(x => x.FamilyName).Distinct().ElementAt(index2);
                    FamilySymbol2 = Symbols.Where(x => x.FamilyName.Equals(familySymbolName2)).First();
                }
            }

            List<XYZ> pipeHangPts = new List<XYZ>();
            try
            {

                OutsideDiam = Element.LookupParameter("Outside Diameter").AsDouble();
                PipeDiam = (OutsideDiam) + (2 * Element.LookupParameter("Insulation Thickness").AsDouble());
            }
            catch
            {
                PipeDiam = (Width / 304.8) + (2 * Element.LookupParameter("Insulation Thickness").AsDouble());

            }
            HangerDiameter = -1;
            foreach (double hangerDiam in Diameters)
            {
                if (PipeDiam <= (hangerDiam / 304.8))
                {
                    HangerDiameter = hangerDiam / 304.8;
                    break;
                }
            }
            if (HangerDiameter == -1) { HangerDiameter = 402 / 304.8; }
            levelId = Element.LookupParameter("Reference Level").AsElementId();
            pipeElevation = ElementCurve.GetEndPoint(0).Z - Element.LookupParameter("Insulation Thickness").AsDouble();
            midElevStart = Element.LookupParameter("Start Middle Elevation").AsDouble();
            midElevEnd = Element.LookupParameter("End Middle Elevation").AsDouble();
            slope = Element.LookupParameter("Slope").AsDouble();
            if (ElementCurve.Length > Negligible && ElementCurve.Length <= 5 * Offset)
            {
                XYZ midPt = ElementCurve.Evaluate(0.50, true);
                if (!pipeHangPts.Contains(midPt))
                {
                    pipeHangPts.Add(midPt);
                    double rod = GetRod(midPt);
                    if (rod != 0) Supports.Add(new Support(midPt, rod + 2995 / 304.8));
                }
            }
            else if (ElementCurve.Length <= Spacing && ElementCurve.Length > 5 * Offset)
            {
                if (!pipeHangPts.Contains(Ps))
                {
                    pipeHangPts.Add(Ps);
                    double rod = GetRod(Ps);
                    if (rod != 0) Supports.Add(new Support(Ps, rod + 2995 / 304.8));

                }
                if (!pipeHangPts.Contains(Pe))
                {
                    pipeHangPts.Add(Pe);
                    double rod = GetRod(Pe);
                    if (rod != 0) Supports.Add(new Support(Pe, rod + 2995 / 304.8));

                }
            }
            else if (ElementCurve.Length > Spacing)
            {
                if (!pipeHangPts.Contains(Ps))
                {
                    pipeHangPts.Add(Ps);
                    double rod = GetRod(Ps);
                    if (rod != 0) Supports.Add(new Support(Ps, rod + 2995 / 304.8));
                }
                double n = Math.Ceiling(hangCurve.Length / Spacing) - 1;
                XYZ prev = Ps;
                double Ns = (hangCurve.Length / (n + 1));
                for (int i = 0; i < n; i++)
                {
                    XYZ point = prev.Add(Ns * pipeDirection);
                    if (!pipeHangPts.Contains(point))
                    {
                        pipeHangPts.Add(point);
                        double rod = GetRod(point);
                        if (rod != 0) Supports.Add(new Support(point, rod + 2995 / 304.8));
                    }
                    prev = point;
                }
                if (!pipeHangPts.Contains(Pe))
                {
                    pipeHangPts.Add(Pe);
                    double rod = GetRod(Pe);
                    if (rod != 0) Supports.Add(new Support(Pe, rod + 2995 / 304.8));

                }
            }
            if (Supports.Count > 0) isValid = true;
        }

        public void Plant()
        {
            if (!isValid) return;
            foreach (Support support in Supports)
            {
                if (isFireFighting)
                {
                    FamilySymbol.Activate();
                    FamilyInstance hangerFamInst = Document.Create
                        .NewFamilyInstance(support.point, FamilySymbol, Perpendicular, Document.GetElement(levelId), StructuralType.NonStructural);
                    double q = 1;
                    if (midElevEnd < midElevStart)
                    {
                        q = -1;
                    }
                    double elev = pipeElevation + (q * slope * support.point.DistanceTo(startPt));
                    double pipeOffsetFromHost = elev - (3000 / 304.8);
                    if (hangerFamInst == null) { return; }
                    hangerFamInst.LookupParameter("Diameter").Set(PipeDiam);
                    hangerFamInst.LookupParameter("Offset from Host").Set(pipeOffsetFromHost);
                    hangerFamInst.LookupParameter("AnchorElevation").Set(support.rod);
                }
                else
                {
                    Reference reference = null;
                    if (Math.Round(Math.Abs(((PlanarFace)LowerFace).FaceNormal.Z), 3) == 1)
                    {
                        if (LowerFace.Reference.CreateLinkReference(DocumentRLI) != null)
                        {
                            reference = LowerFace.Reference.CreateLinkReference(DocumentRLI);
                        }
                    }
                    else // Ramp
                    {
                        FilteredElementCollector airs = new FilteredElementCollector(Document).OfClass(typeof(FamilySymbol));
                        FamilySymbol fsair = null;
                        foreach (FamilySymbol fsa in airs)
                        {
                            if (fsa.Name == "Air") { fsair = fsa; fsair.Activate(); }
                        }
                        FamilyInstance air = null;
                        air = Document.Create.NewFamilyInstance(support.point, fsair, StructuralType.NonStructural);
                        Document.Regenerate();
                        Face faceair = GetSolid(air).Faces.get_Item(0);
                        string refnew = air.UniqueId + ":0:INSTANCE:" + faceair.Reference.ConvertToStableRepresentation(Document);
                        reference = Reference.ParseFromStableRepresentation(Document, refnew);
                    }
                    try
                    {
                        FamilySymbol fs = null;
                        if (Width > (202))
                        {
                            fs = FamilySymbol2;
                            if (fs != null)
                            {
                                fs.Activate();
                            }
                            else
                            {
                                fs = FamilySymbol;
                                fs.Activate();
                            }
                        }
                        else
                        {
                            fs = FamilySymbol;
                            fs.Activate();
                        }
                        FamilyInstance pang = Document.Create.NewFamilyInstance(reference, support.point, pipeDirection, fs);
                        pang.LookupParameter("Schedule Level").Set(levelId);
                        Line ll = Line.CreateUnbound(support.point, XYZ.BasisZ);
                        double rr = pang.HandOrientation.AngleOnPlaneTo(pipeDirection, XYZ.BasisZ);
                        IntersectionResultArray iraa = new IntersectionResultArray();
                        SetComparisonResult scr = LowerFace.Intersect(ll, out iraa);
                        if (iraa != null && !iraa.IsEmpty)
                        {
                            Curve cv = Line.CreateBound(support.point, iraa.get_Item(0).XYZPoint);
                            pang.LookupParameter("Pipe_distance").Set(cv.Length - (0.5 * PipeDiam));
                        }
                        pang.LookupParameter("Pipe Outer Diameter").Set(PipeDiam);
                    }
                    catch
                    {
                    }
                }
            }
        }




    }
}