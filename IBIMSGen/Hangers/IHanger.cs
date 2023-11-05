using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace IBIMSGen.Hangers
{
    public interface IHanger
    {
        Document Document { get; }
        Solid Solid { get; }
        Element Element { get; }
        List<List<Dictionary<string, double>>> Dimensions { get; }
        double Up { get; }
        double Down { get; }
        List<FamilySymbol> Symbols { get; }
        double Negligible { get; }
        double Offset { get; }
        List<Support> Supports { get; }
        FamilySymbol FamilySymbol { get; }
        XYZ Perpendicular { get; }
        double Width { get; }
        bool isValid { get; }
        double Spacing { get; }
        RevitLinkInstance LinkInstance { get; }
        void Process();
        void Plant();
        double GetRod(XYZ point);
        List<XYZ> DecOrder(List<XYZ> points, Curve curve);
        int GetSystemRank(string name);
        double GetSysSpacing(List<Dictionary<string, double>> dimensions, double diameter);
    }
}
