using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IBIMSGen.Penetration
{
    public class PenetratingElement
    {
        public Element element;
        public double width, height, insulationThickness;
        public WorksetId worksetId;
        public Curve axis;
        public FamilySymbol familySymbol;
        public XYZ sleeveDir;
        RevitLinkInstance rli;
        public PenetratingElement(Element element, WorksetId worksetId,
            double width, double height, double insulationThickness, Curve axis,
            FamilySymbol familySymbol)
        {
            this.element = element;
            this.worksetId = worksetId;
            this.width = width;
            this.height = height;
            this.insulationThickness = insulationThickness;
            this.axis = axis;
            this.familySymbol = familySymbol;
        }
        public PenetratingElement(Element element, RevitLinkInstance rli)
        {
            this.element = element;
            this.rli = rli;
        }
    }

}
