using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IBIMSGen.Penetration
{
    //selection filter class
    public class selectionFilter : ISelectionFilter
    {
        public Document doc;
        public bool AllowElement(Element elem)
        {
            doc = ((RevitLinkInstance)elem).GetLinkDocument();
            return true;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            Element ee = doc.GetElement(reference.LinkedElementId);
            if (ee is Pipe || ee is Duct || ee is Conduit || ee is CableTray || ee is Floor || ee is Wall || ee.Category.Name == "Structural Framing" || ee.Category.Name == "Structural Columns" || ee is FamilyInstance)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
