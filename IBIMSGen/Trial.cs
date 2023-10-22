using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace IBIMSGen
{
    [TransactionAttribute(TransactionMode.Manual)]
    internal class Trial : IExternalCommand
    {
        UIDocument uidoc;
        Document doc;
        StringBuilder sb;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            sb = new StringBuilder();

            TaskDialog.Show("Info", sb.ToString());
            return Result.Succeeded;
        }

        void getGrid()
        {
            Element selectedElement = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element));
            sb.AppendLine(selectedElement.Name);
            var material = doc.GetElement(selectedElement.GetMaterialIds(false).ElementAt(1)) as Material;
            sb.AppendLine(material.Name);

            FilteredElementCollector fec = new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement));
            var patternElementId = material.SurfaceForegroundPatternId;
            var patternElement = fec.Where(x => x.Id == patternElementId).FirstOrDefault();
            var pattern = patternElement as FillPatternElement;
            sb.AppendLine(pattern?.Name);
            var fillPattern = pattern.GetFillPattern().GetFillGrid(1);
            var patOrigin = fillPattern.Origin;
            sb.AppendLine(patOrigin.ToString());
            var offset = fillPattern.Offset;
            sb.AppendLine(offset.ToString());
            var segments=fillPattern.GetSegments();
            sb.AppendLine(segments.Count.ToString());
            try
            {
            }
            catch
            {

            }
        }
    }
}
