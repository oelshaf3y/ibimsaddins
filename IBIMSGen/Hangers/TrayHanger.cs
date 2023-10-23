using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace IBIMSGen.Hangers
{
    internal class TrayHanger
    {
        public Element tray;
        public double trayWidth;
        public ElementId levelId;
        public double trayElevation;
        public List<XYZ> hangPts;
        public XYZ trayDir;
        public XYZ trayPerpendicular;

        public TrayHanger(Element tray, double trayWidth, ElementId levelId, double trayElevation, List<XYZ> hangPts, XYZ trayDir, XYZ trayPerpendicular)
        {
            this.tray = tray;
            this.trayWidth = trayWidth;
            this.levelId = levelId;
            this.trayElevation = trayElevation;
            this.hangPts = hangPts;
            this.trayDir = trayDir;
            this.trayPerpendicular = trayPerpendicular;
        }
    }
}