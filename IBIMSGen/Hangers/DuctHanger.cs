using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using System;
using System.Collections.Generic;
using System.Linq;

namespace IBIMSGen.Hangers
{
    internal class DuctHanger : IHanger
    {
        public Document Document { get; }
        public Solid Solid { get; private set; }
        public Element Element { get; }
        public List<List<Dictionary<string, double>>> Dimensions { get; }
        public List<FamilySymbol> Symbols { get; }
        public double Negligible { get; }
        public double Offset { get; }
        public double Up { get; }
        public double Down { get; }
        public bool isValid { get; private set; }
        public List<Element> FitsInRange { get; private set; }
        public XYZ Perpendicular { get; private set; }
        public double Width { get; private set; }
        public double Height { get; private set; }
        public List<Support> Supports { get; private set; } = new List<Support>();
        public FamilySymbol FamilySymbol { get; private set; }
        public double InsulationThick { get; private set; }
        public double BotElevation { get; private set; }
        public double Spacing { get; private set; } = 0;

        public RevitLinkInstance LinkInstance { get; }

        public DuctHanger(Document document, Solid solid, Element element, List<List<Dictionary<string, double>>> dimensions, List<FamilySymbol> symbols,
            double negligible, double offset, double floorUp, double floorDown, List<Element> fitsInRange, RevitLinkInstance RLI = null)
        {
            Document = document;
            Solid = solid;
            Element = element;
            Dimensions = dimensions;
            Symbols = symbols;
            Negligible = negligible;
            Offset = offset;
            Up = floorUp;
            Down = floorDown;
            FitsInRange = fitsInRange;
            LinkInstance = RLI;
            Process();
        }
        public void Process()
        {
            List<Support> _supports = new List<Support>();

            Curve ductCurve = ((LocationCurve)Element.Location).Curve;
            double minx = Math.Min(ductCurve.GetEndPoint(0).X, ductCurve.GetEndPoint(1).X) - 5;
            double miny = Math.Min(ductCurve.GetEndPoint(0).Y, ductCurve.GetEndPoint(1).Y) - 5;
            double maxx = Math.Max(ductCurve.GetEndPoint(0).X, ductCurve.GetEndPoint(1).X) + 5;
            double maxy = Math.Max(ductCurve.GetEndPoint(0).Y, ductCurve.GetEndPoint(1).Y) + 5;
            Boundary region = new Boundary(minx, maxy, maxx, miny, Up, Down);
            XYZ ductOrigin = ductCurve.Evaluate(0.5, true);
            Width = 0;
            List<XYZ> pts = new List<XYZ>();
            try { Width = Element.LookupParameter("Width").AsDouble(); }
            catch { Width = Element.LookupParameter("Diameter").AsDouble(); }
            if (Dimensions[0].Count == 1)
            {
                Spacing = Dimensions[0][0]["spacing"] / 304.8;
                if (Spacing == 0) { isValid = false; return; }
                string familySymbolName = Symbols.Select(x => x.FamilyName).Distinct().ElementAt(Convert.ToInt32(Dimensions[0][0]["family"]));
                FamilySymbol = Symbols.Where(x => x.FamilyName.Equals(familySymbolName)).First();
            }
            else
            {
                if (Dimensions[0].Where(x => x["from"] < Width * 304.8 && Width * 304.8 <= x["to"]).Any())
                {
                    Spacing = Dimensions[0].Where(x => x["from"] < Width * 304.8 && Width * 304.8 <= x["to"]).First()["spacing"] / 304.8;
                    if (Spacing == 0) { isValid = false; return; }
                    int index = Convert.ToInt32(Dimensions[0][0]["family"]);
                    if (index < 0) return;
                    string familySymbolName = Symbols.Select(x => x.FamilyName).Distinct().ElementAt(index);
                    FamilySymbol = Symbols.Where(x => x.FamilyName.Equals(familySymbolName)).First();
                }
            }
            ElementId levelId = Element.LookupParameter("Reference Level").AsElementId();
            Height = 0;
            try { Height = Element.LookupParameter("Height").AsDouble(); }
            catch { Height = Element.LookupParameter("Diameter").AsDouble(); }
            double botElevationParam;
            try
            {
                botElevationParam = Element.LookupParameter("Bottom Elevation").AsDouble();
            }
            catch
            {
                botElevationParam = Element.LookupParameter("Lower End Bottom Elevation").AsDouble();
            }
            BotElevation = botElevationParam + ((Level)Document.GetElement(levelId)).Elevation;
            InsulationThick = Element.LookupParameter("Insulation Thickness").AsDouble();
            XYZ ductDir = ((Line)ductCurve).Direction.Normalize();
            Perpendicular = new XYZ(-ductDir.Y, ductDir.X, ductDir.Z);
            XYZ ductMidPt = ductCurve.Evaluate(0.5, true);

            #region Duct Fitting
            // getting the duct fitting points that intersects the curve line.
            // then offsets the hangers by a margin that insures the hangers will never intersect the fitting.

            List<Element> nearFits = FitsInRange.Where(x => region.contains(x)).ToList();
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
                double radius = Math.Sqrt(Math.Pow(Width, 2) + Math.Pow(width1Param.AsDouble(), 2)) * 0.50;
                XYZ fittingLocPt = ((LocationPoint)fitting.Location).Point;
                double margin = 0.50 * (Math.Tan(angle) * takeOffLength);
                XYZ center = new XYZ(fittingLocPt.X, fittingLocPt.Y, ductMidPt.Z).Add(-margin * fittingFI.FacingOrientation);
                Curve circ = Ellipse.CreateCurve(center, radius, radius, ductDir, Perpendicular, 0, 2 * Math.PI * radius);
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
            List<XYZ> fittingPts = DecOrder(pts, ductCurve);
            List<XYZ> ductCurvePts = DecOrder(new List<XYZ>() { P0, Pf }, ductCurve);
            XYZ dir = Line.CreateBound(ductCurvePts[0], ductCurvePts[1]).Direction.Normalize();
            XYZ Ps = ductCurvePts[0].Add(Offset * dir);
            XYZ Pe = ductCurvePts[1].Add(-Offset * dir);
            Curve cc = null;
            try
            {
                cc = Line.CreateBound(Ps, Pe);
            }
            catch { isValid = false; return; }
            if (fittingPts.Count == 0) // No ductfittings
            {
                if (ductCurve.Length > Negligible && ductCurve.Length <= 4 * Offset)
                {
                    double rod = GetRod(ductMidPt);
                    XYZ PP = new XYZ(ductMidPt.X, ductMidPt.Y, ductMidPt.Z - InsulationThick - (Height / 2));

                    if (rod != 0) _supports.Add(new Support(PP, rod));
                }
                else if (ductCurve.Length <= Spacing && ductCurve.Length > 4 * Offset)
                {
                    double rod1 = GetRod(Ps);
                    XYZ ps = new XYZ(Ps.X, Ps.Y, Ps.Z - InsulationThick - (Height / 2));
                    if (rod1 != 0) _supports.Add(new Support(ps, rod1));

                    double rod2 = GetRod(Pe);
                    XYZ pe = new XYZ(Pe.X, Pe.Y, Pe.Z - InsulationThick - (Height / 2));
                    if (rod2 != 0) _supports.Add(new Support(pe, rod2));
                }
                else if (ductCurve.Length > Spacing) // collect point of hangers
                {
                    double rod = GetRod(Ps);
                    XYZ ps = new XYZ(Ps.X, Ps.Y, Ps.Z - InsulationThick - (Height / 2));
                    if (rod != 0) _supports.Add(new Support(ps, rod));
                    double n = Math.Ceiling(cc.Length / Spacing) - 1;
                    XYZ prevPt = Ps;
                    for (int i = 0; i < n; i++)
                    {
                        XYZ point = prevPt.Add(Spacing * dir);
                        double ROD = GetRod(point);
                        XYZ p = new XYZ(point.X, point.Y, point.Z - InsulationThick - (Height / 2));
                        if (ROD != 0) _supports.Add(new Support(p, ROD));
                        prevPt = point;
                    }
                    double rod2 = GetRod(Pe);
                    XYZ pe = new XYZ(Pe.X, Pe.Y, Pe.Z - InsulationThick - (Height / 2));
                    if (rod2 != 0) _supports.Add(new Support(pe, rod2));
                }
            }
            else       // With ductfittings
            {
                if (ductCurve.Length > Negligible && ductCurve.Length <= 4 * Offset)
                {
                    List<XYZ> fitPtsOrdered = DecOrder(new List<XYZ>() { fittingPts[0], fittingPts[1], ductMidPt }, ductCurve);
                    if (fitPtsOrdered.IndexOf(ductMidPt) == 1)
                    {
                        double rod = GetRod(fitPtsOrdered[2]);
                        XYZ point = new XYZ(fitPtsOrdered[2].X, fitPtsOrdered[2].Y, fitPtsOrdered[2].Z - InsulationThick - (Height / 2));
                        if (rod != 0) _supports.Add(new Support(point, rod)); // farther point
                    }
                    else
                    {
                        double rod = GetRod(ductMidPt);
                        XYZ point = new XYZ(ductMidPt.X, ductMidPt.Y, ductMidPt.Z - InsulationThick - (Height / 2));
                        if (rod != 0) _supports.Add(new Support(point, rod)); // duct center
                    }
                }

                else if (ductCurve.Length <= Spacing && ductCurve.Length > 4 * Offset)
                {

                    List<XYZ> fitPtsOrdered = DecOrder(new List<XYZ>() { fittingPts[0], fittingPts[1], Ps }, ductCurve);
                    if (fitPtsOrdered.IndexOf(Ps) == 1)
                    {
                        double rod = GetRod(fitPtsOrdered[2]);
                        XYZ point = new XYZ(fitPtsOrdered[2].X, fitPtsOrdered[2].Y, fitPtsOrdered[2].Z - InsulationThick - (Height / 2));
                        if (rod != 0) _supports.Add(new Support(point, rod));
                    }
                    else
                    {
                        double rod = GetRod(Ps);
                        XYZ point = new XYZ(Ps.X, Ps.Y, Ps.Z - InsulationThick - (Height / 2));
                        if (rod != 0) _supports.Add(new Support(point, rod));
                    }
                    List<XYZ> ptss = new List<XYZ>() { Ps, Pe };
                    for (int i = 0; i < 2; i++)
                    {
                        for (int j = 0; j < fittingPts.Count; j += 2)
                        {
                            List<XYZ> pso = DecOrder(new List<XYZ>() { fittingPts[j], fittingPts[j + 1], ptss[i] }, ductCurve);
                            if (pso.IndexOf(ptss[i]) == 1) //Between Case
                            {
                                if (i == 0)
                                {
                                    double rod = GetRod(pso[2]);
                                    XYZ point = new XYZ(pso[2].X, pso[2].Y, pso[2].Z - InsulationThick - (Height / 2));
                                    if (rod != 0) _supports.Add(new Support(point, rod));
                                }
                                else
                                {
                                    double rod = GetRod(pso[0]);
                                    XYZ point = new XYZ(pso[0].X, pso[0].Y, pso[0].Z - InsulationThick - (Height / 2));
                                    if (rod != 0) _supports.Add(new Support(point, rod));
                                }
                                break;
                            }
                        }
                    }
                    double rod2 = GetRod(Pe);
                    XYZ pe = new XYZ(Pe.X, Pe.Y, Pe.Z - InsulationThick - (Height / 2));
                    if (rod2 != 0) _supports.Add(new Support(pe, rod2));
                }
                else if (ductCurve.Length > Spacing)
                {
                    List<XYZ> ps1o = DecOrder(new List<XYZ>() { fittingPts[0], fittingPts[1], Ps }, ductCurve);
                    if (ps1o.IndexOf(Ps) == 1)
                    {
                        double rod = GetRod(ps1o[2]);
                        XYZ point = new XYZ(ps1o[2].X, ps1o[2].Y, ps1o[2].Z - InsulationThick - (Height / 2));
                        if (rod != 0) _supports.Add(new Support(point, rod));
                    }
                    else
                    {
                        double rod = GetRod(Ps);
                        XYZ ps = new XYZ(Ps.X, Ps.Y, Ps.Z - InsulationThick - (Height / 2));
                        if (rod != 0) _supports.Add(new Support(ps, rod));
                    }
                    double n = Math.Ceiling(cc.Length / Spacing) - 1;
                    XYZ prevPt = Ps;
                    for (int i = 0; i < n; i++)
                    {
                        XYZ point = prevPt.Add(Spacing * dir);
                        for (int j = 0; j < fittingPts.Count; j += 2)
                        {
                            List<XYZ> ps3o = DecOrder(new List<XYZ>() { fittingPts[j], fittingPts[j + 1], point }, ductCurve);
                            if (ps3o.IndexOf(point) == 1) //Between Case
                            {
                                point = ps3o[0];
                                break;
                            }
                        }
                        double rod = GetRod(point);
                        XYZ P = new XYZ(point.X, point.Y, point.Z - InsulationThick - (Height / 2));
                        if (rod != 0) _supports.Add(new Support(point, rod));
                        prevPt = point;
                    }
                }
            }
            foreach (Support sup in _supports)
            {
                Support found = null;
                found = Supports.Where(x => x.point.IsAlmostEqualTo(sup.point))?.FirstOrDefault();
                if (found != null)
                {
                    if (found.rod > sup.rod) found.rod = sup.rod;
                }
                else
                {
                    Supports.Add(sup);
                }
            }
            if (Supports.Count > 0) isValid = true;
        }
        public List<XYZ> DecOrder(List<XYZ> points, Curve curve)
        {
            if (Math.Round(((Line)curve).Direction.Normalize().Y, 3) == 0)
            {
                return points.OrderByDescending(a => a.X).ToList();
            }
            else
            {
                return points.OrderByDescending(a => a.Y).ToList();
            }
        }
        public void Plant()
        {
            if (!isValid) return;
            FamilySymbol.Activate();
            foreach (Support support in Supports)
            {
                FamilyInstance hang = Document.Create.NewFamilyInstance(support.point, FamilySymbol, Perpendicular, Element, StructuralType.NonStructural);
                hang.LookupParameter("Width").Set(Width + (2 * InsulationThick) + 16 / 304.8);
                double Z = BotElevation - InsulationThick - hang.LookupParameter("Elevation from Level").AsDouble();
                hang.Location.Move(new XYZ(0, 0, Z));
                support.rod += InsulationThick + (Height / 2);
                hang.LookupParameter("ROD 1").Set(support.rod);
                hang.LookupParameter("ROD 2").Set(support.rod);
            }
        }
        public double GetRod(XYZ point)
        {
            Line tempLine = Line.CreateUnbound(point, XYZ.BasisZ);
            Face lower = Solid.Faces.get_Item(0);
            //Face upper = Solid.Faces.get_Item(1);
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
        public int GetSystemRank(string name) => throw new NotImplementedException();
        public double GetSysSpacing(List<Dictionary<string, double>> dimensions, double diameter) => throw new NotImplementedException();
    }
}