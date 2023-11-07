using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;

namespace IBIMSGen.Hangers
{
    internal class TrayHanger : Hanger
    {
        public RevitLinkInstance DocumentRLI { get; }
        public ElementId LevelId { get; private set; }
        public XYZ trayDir { get; private set; }
        public double botElevation { get; private set; }

        public TrayHanger(Document document, Solid solid, Element element, List<List<Dictionary<string, double>>> dimensions, double up, double down,
            List<FamilySymbol> symbols, double negligible, double offset, RevitLinkInstance linkInstance,
            QuadTree allDuctsTree, QuadTree allPipesTree, QuadTree allTraysTree, RevitLinkInstance trayRLI = null)
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
            AllDuctsTree = allDuctsTree;
            AllPipesTree = allPipesTree;
            AllTraysTree = allTraysTree;
            Process();
        }



        public void Process()
        {
            FamilySymbol = null;
            ElementCurve = ((LocationCurve)Element.Location).Curve;
            GetRegion();
            double trayOffset = 500 / 304.80;
            trayDir = ((Line)ElementCurve).Direction.Normalize();
            Perpendicular = new XYZ(-trayDir.Y, trayDir.X, trayDir.Z);
            XYZ P0 = ElementCurve.Evaluate(0, true);
            XYZ Pf = ElementCurve.Evaluate(1, true);
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
            if (ElementCurve.Length > Negligible && ElementCurve.Length <= trayOffset)
            {
                XYZ midPt = ElementCurve.Evaluate(0.50, true);
                if (!trayHangPts.Contains(midPt))
                {
                    trayHangPts.Add(midPt);
                    double rod = GetRod(midPt);
                    XYZ PP = new XYZ(midPt.X, midPt.Y, elevation);
                    if (rod != 0) Supports.Add(new Support(PP, rod));
                }
            }
            else if (ElementCurve.Length <= Spacing && ElementCurve.Length > trayOffset)
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
            else if (ElementCurve.Length > Spacing)
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


    }
}