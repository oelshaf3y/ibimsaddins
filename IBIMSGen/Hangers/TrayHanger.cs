using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace IBIMSGen.Hangers
{
    internal class TrayHanger : IHanger
    {
        public Solid Solid { get; }
        public Element Element { get; }
        public double Up { get; }
        public double Down { get; }
        public double Width { get; private set; }
        public List<Support> Supports { get; private set; }
        public XYZ Perpendicular { get; private set; }
        public FamilySymbol FamilySymbol { get; private set; }
        public Document Document { get; }
        public List<List<Dictionary<string, double>>> Dimensions { get; }
        public List<FamilySymbol> Symbols { get; }
        public double Negligible { get; }
        public double Offset { get; }
        public bool isValid { get; private set; }

        public double Spacing { get; private set; }

        public RevitLinkInstance LinkInstance { get; }

        public ElementId levelId;
        public XYZ trayDir;
        public TrayHanger(Element tray, double trayWidth, ElementId levelId, List<Support> supports, XYZ trayDir, XYZ trayPerpendicular, FamilySymbol familySymbol)
        {
            Element = tray;
            Width = trayWidth;
            this.levelId = levelId;
            Supports = supports;
            this.trayDir = trayDir;
            Perpendicular = trayPerpendicular;
            FamilySymbol = familySymbol;
        }

        public void Process()
        {
            throw new System.NotImplementedException();
        }

        public void Plant()
        {
            throw new System.NotImplementedException();
        }

        public double GetRod(XYZ point)
        {
            throw new System.NotImplementedException();
        }

        public List<XYZ> DecOrder(List<XYZ> points, Curve cuurve)
        {
            throw new System.NotImplementedException();
        }

        public int GetSystemRank(string name)
        {
            throw new System.NotImplementedException();
        }

        public double GetSysSpacing(List<Dictionary<string, double>> dimensions, double diameter)
        {
            throw new System.NotImplementedException();
        }
    }
}