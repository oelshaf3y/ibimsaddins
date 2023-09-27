using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;

namespace IBIMSGen
{
    public partial class Penetration
    {
        public class strSelectionFilter : ISelectionFilter
        {
            Document doc;
            public bool AllowElement(Element elem)
            {
                if(elem is RevitLinkInstance)
                {
                    doc = ((RevitLinkInstance)elem).GetLinkDocument();

                    return true;
                }
                return false;
            }

            public bool AllowReference(Reference reference, XYZ position)
            {
                if(doc == null) { TaskDialog.Show("err", "null"); return false; }
                
                Element elem = doc.GetElement(reference.LinkedElementId);

                if (elem is Wall)
                {
                    return true;
                }
                return false;
            }
        }
    }
}
