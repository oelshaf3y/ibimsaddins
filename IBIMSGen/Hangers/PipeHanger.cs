using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace IBIMSGen.Hangers
{
    internal class PipeHanger
    {
        public Element pipe;
        public XYZ startPt;
        public bool isFireFighting;
        public double hangerDiameter;
        public ElementId levelId;
        public double pipeElevation;
        public double midElevStart;
        public double midElevEnd;
        public double slope;
        public List<Support> supports;
        public XYZ pipeDirection;
        public XYZ pipePerpendicular;
        public Face lowerFace;

        public PipeHanger(Element pipe, XYZ startPt, bool isFireFighting, double hangerDiameter, ElementId levelId, double pipeElevation, double midElevStart, double midElevEnd, double slope, List<Support> supports, XYZ pipeDirection, XYZ pipePerpendicular, Face lowerFace)
        {
            this.pipe = pipe;
            this.startPt = startPt;
            this.isFireFighting = isFireFighting;
            this.hangerDiameter = hangerDiameter;
            this.levelId = levelId;
            this.pipeElevation = pipeElevation;
            this.midElevStart = midElevStart;
            this.midElevEnd = midElevEnd;
            this.slope = slope;
            this.supports = supports;
            this.pipeDirection = pipeDirection;
            this.pipePerpendicular = pipePerpendicular;
            this.lowerFace = lowerFace;
        }
    }
}