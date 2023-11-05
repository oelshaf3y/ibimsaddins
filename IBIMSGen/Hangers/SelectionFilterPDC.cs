using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI.Selection;

namespace IBIMSGen.Hangers
{
    public class SelectionFilterPDC : ISelectionFilter
    {
        public Document doc;
        public RevitLinkInstance RLI { get; }
        public SelectionFilterPDC(RevitLinkInstance rli = null)
        {
            RLI = rli;
        }
        public bool AllowElement(Element e)
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
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            Element ee = doc.GetElement(reference.LinkedElementId);
            if (ee.Category != null)
            {
                if (ee is Pipe || ee is Duct || ee is CableTray)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return false;
            //throw new NotImplementedException();
        }
    }

}


