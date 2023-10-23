using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace IBIMSGen.Hangers
{
    internal class DuctHanger
    {
        public Element duct;
        public XYZ ductPerpendicular;
        public double ductWidth;
        public double ductHeight;
        public double insoThick;
        public double botElevation;
        public List<XYZ> hangPts;

        public DuctHanger(Element duct, XYZ ductPerpendicular, double ductWidth, double ductHeight, double insoThick, double botElevation, List<XYZ> hangPts)
        {
            this.duct = duct;
            this.ductPerpendicular = ductPerpendicular;
            this.ductWidth = ductWidth;
            this.ductHeight = ductHeight;
            this.insoThick = insoThick;
            this.botElevation = botElevation;
            this.hangPts = hangPts;
        }
    }
}