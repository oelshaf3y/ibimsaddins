using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;

namespace IBIMSGen.Hangers
{
    public class SelectionFilterPDC : ISelectionFilter
    {
        public bool AllowElement(Element e)
        {
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
            throw new NotImplementedException();
        }
    }

}


