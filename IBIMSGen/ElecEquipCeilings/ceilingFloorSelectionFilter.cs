using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace IBIMSGen.ElecEquipCeilings
{
    public class ceilingFloorSelectionFilter : ISelectionFilter
    {
        Document link;
        public ceilingFloorSelectionFilter()
        {
            link = null;
        }

        public bool AllowElement(Element elem)
        {
            if (elem is RevitLinkInstance)
            {

                link = ((RevitLinkInstance)elem).GetLinkDocument();
                return true;
            }
            return (elem is Floor || elem is Ceiling);
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return link == null;
            Element element = link.GetElement(reference.LinkedElementId);
            return (element is Floor || element is Ceiling);
        }
    }
}
