using Autodesk.Revit.DB;

namespace IBIMSGen.Hangers
{
    internal class Support
    {
        public XYZ point { get; set; }
        public double rod { get; set; }
        public Support(XYZ point, double rod)
        {
            this.point = point;
            this.rod = rod;
        }
    }
}