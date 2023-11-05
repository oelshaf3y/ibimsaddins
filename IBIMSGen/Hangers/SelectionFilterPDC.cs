using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI.Selection;
using System;

namespace IBIMSGen.Hangers
{
    public class SelectionFilterPDC : ISelectionFilter
    {
        public Document doc;
        public RevitLinkInstance RLI { get; }
        public bool isHost;
        public SelectionFilterPDC(RevitLinkInstance rli = null, bool isHost = false)
        {
            RLI = rli;
            this.isHost = isHost;
        }
        public bool AllowElement(Element e)
        {
            if (this.isHost)
            {

                doc = ((RevitLinkInstance)e).GetLinkDocument();
                if (doc != null) return true;

                if (e.Category != null)
                {
                    if (e is Floor)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            else
            {

                if (RLI != null && ((RevitLinkInstance)e).Name != RLI.Name)
                {
                    doc = ((RevitLinkInstance)e).GetLinkDocument();
                    if (doc != null) return true;
                }
                if (e.Category != null)
                {
                    if (e is Pipe || e is Duct || e is CableTray)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            Element ee = doc.GetElement(reference.LinkedElementId);
            if (isHost)
            {
                if (ee.Category != null) return ee is Floor;
            }
            else
            {
                if (ee.Category != null)
                {
                    return (ee is Pipe || ee is Duct || ee is CableTray);
                }
            }
            return false;
            //throw new NotImplementedException();
        }
    }

}


