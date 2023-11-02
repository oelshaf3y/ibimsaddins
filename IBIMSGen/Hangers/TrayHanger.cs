using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace IBIMSGen.Hangers
{
    internal class TrayHanger
    {
        public Element tray;
        public double trayWidth;
        public ElementId levelId;
        public List<Support> supports;
        public XYZ trayDir;
        public XYZ trayPerpendicular;
        public FamilySymbol familySymbol;
        public TrayHanger(Element tray, double trayWidth, ElementId levelId,List<Support> supports, XYZ trayDir, XYZ trayPerpendicular, FamilySymbol familySymbol)
        {
            this.tray = tray;
            this.trayWidth = trayWidth;
            this.levelId = levelId;
            this.supports = supports;
            this.trayDir = trayDir;
            this.trayPerpendicular = trayPerpendicular;
            this.familySymbol = familySymbol;
        }
    }
}