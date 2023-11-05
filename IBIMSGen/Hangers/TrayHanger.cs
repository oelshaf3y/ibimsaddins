using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;

namespace IBIMSGen.Hangers
{
    internal class TrayHanger : IHanger
    {
        public Document Document { get; }
        public Solid Solid { get; }
        public Element Element { get; }
        public List<List<Dictionary<string, double>>> Dimensions { get; }
        public double Up { get; }
        public double Down { get; }
        public List<FamilySymbol> Symbols { get; }
        public double Negligible { get; }
        public double Offset { get; }
        public RevitLinkInstance DocumentRLI { get; }
        public RevitLinkInstance LinkInstance { get; }

        public List<Support> Supports { get; private set; } = new List<Support>();
        public FamilySymbol FamilySymbol { get; private set; }
        public XYZ Perpendicular { get; private set; }
        public double Width { get; private set; }
        public bool isValid { get; private set; } = false;
        public double Spacing { get; private set; } = 0;
        public ElementId LevelId { get; private set; }
        public XYZ trayDir { get; private set; }
        public double botElevation { get; private set; }

        public TrayHanger(Document document, Solid solid, Element element, List<List<Dictionary<string, double>>> dimensions, double up, double down,
            List<FamilySymbol> symbols, double negligible, double offset, RevitLinkInstance linkInstance, RevitLinkInstance trayRLI = null)
        {
            Document = document;
            Solid = solid;
            Element = element;
            Dimensions = dimensions;
            Up = up;
            Down = down;
            Symbols = symbols;
            Negligible = negligible;
            Offset = offset;
            DocumentRLI = linkInstance;
            LinkInstance = trayRLI;
            Process();
        }



        public void Process()
        {
            FamilySymbol = null;
            Curve trayCurve = ((LocationCurve)Element.Location).Curve;
            double trayOffset = 500 / 304.80;
            trayDir = ((Line)trayCurve).Direction.Normalize();
            Perpendicular = new XYZ(-trayDir.Y, trayDir.X, trayDir.Z);
            XYZ P0 = trayCurve.Evaluate(0, true);
            XYZ Pf = trayCurve.Evaluate(1, true);
            XYZ Ps = P0.Add(Offset * trayDir);
            XYZ Pe = Pf.Add(-Offset * trayDir);
            Width = Element.LookupParameter("Width").AsDouble();
            Spacing = 0;
            //var a = Dimensions[5][0];
            if (Dimensions[5].Count == 0) return;
            if (Dimensions[5].Count == 1)
            {
                Spacing = Dimensions[5][0]["spacing"] / 304.8;
                if (Spacing == 0) return;
                string familySymbolName = Symbols.Select(x => x.FamilyName).Distinct().ElementAt(Convert.ToInt32(Dimensions[5][0]["family"]));
                FamilySymbol = Symbols.Where(x => x.FamilyName.Equals(familySymbolName)).First();
            }
            else
            {
                if (Dimensions[5].Where(x => Width * 304.8 <= x["size"]).Any())
                {
                    Spacing = Dimensions[5].Where(x => Width * 304.8 <= x["size"]).First()["spacing"] / 304.8;
                    if (Spacing == 0) return;
                    string familySymbolName = Symbols.Select(x => x.FamilyName).Distinct().ElementAt(Convert.ToInt32(Dimensions[5][0]["family"]));
                    FamilySymbol = Symbols.Where(x => x.FamilyName.Equals(familySymbolName)).First();
                }
            }
            Curve hangCurve = null;
            try
            {
                hangCurve = Line.CreateBound(Ps, Pe);
            }
            catch { return; }
            List<XYZ> trayHangPts = new List<XYZ>();
            Supports = new List<Support>();
            LevelId = Element.LookupParameter("Reference Level").AsElementId();
            try
            {
                botElevation = Element.LookupParameter("Bottom Elevation").AsDouble();

            }
            catch
            {
                botElevation = Element.LookupParameter("Lower End Bottom Elevation").AsDouble();
            }
            double elevation;
            if (LinkInstance != null)
            {
                elevation = ((Level)LinkInstance.GetLinkDocument().GetElement(LevelId)).Elevation + botElevation;
            }
            else
            {
                elevation = ((Level)Document.GetElement(LevelId)).Elevation + botElevation;
            }
            if (trayCurve.Length > Negligible && trayCurve.Length <= trayOffset)
            {
                XYZ midPt = trayCurve.Evaluate(0.50, true);
                if (!trayHangPts.Contains(midPt))
                {
                    trayHangPts.Add(midPt);
                    double rod = GetRod(midPt);
                    XYZ PP = new XYZ(midPt.X, midPt.Y, elevation);
                    if (rod != 0) Supports.Add(new Support(PP, rod));
                }
            }
            else if (trayCurve.Length <= Spacing && trayCurve.Length > trayOffset)
            {
                if (!trayHangPts.Contains(Ps))
                {
                    trayHangPts.Add(Ps);
                    double rod = GetRod(Ps);
                    XYZ PP = new XYZ(Ps.X, Ps.Y, elevation);
                    if (rod != 0) Supports.Add(new Support(PP, rod));
                }
                if (!trayHangPts.Contains(Pe))
                {
                    trayHangPts.Add(Pe);
                    double rod = GetRod(Pe);
                    XYZ PP = new XYZ(Pe.X, Pe.Y, elevation);
                    if (rod != 0) Supports.Add(new Support(PP, rod));
                }
            }
            else if (trayCurve.Length > Spacing)
            {
                if (!trayHangPts.Contains(Ps))
                {
                    trayHangPts.Add(Ps);
                    double rod = GetRod(Ps);
                    XYZ PP = new XYZ(Ps.X, Ps.Y, elevation);
                    if (rod != 0) Supports.Add(new Support(PP, rod));
                }
                double n = Math.Floor((hangCurve.Length + (100 / 304.8)) / Spacing);
                XYZ prevPt = Ps;
                for (int i = 0; i < n; i++)
                {
                    XYZ point = prevPt.Add(Spacing * trayDir);
                    if (!trayHangPts.Contains(point))
                    {
                        trayHangPts.Add(point);
                        double rod = GetRod(point);
                        XYZ PP = new XYZ(point.X, point.Y, elevation);
                        if (rod != 0) Supports.Add(new Support(PP, rod));
                    }
                    prevPt = point;
                }
            }
            isValid = true;
        }

        public void Plant()
        {
            if (!isValid) return;
            foreach (Support support in Supports)
            {
                FamilySymbol.Activate();
                FamilyInstance hang = Document.Create.NewFamilyInstance(support.point, FamilySymbol, Perpendicular, Document.GetElement(LevelId), StructuralType.NonStructural);
                hang.LookupParameter("Width").Set(Width + 100 / 304.8);
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
    }
}