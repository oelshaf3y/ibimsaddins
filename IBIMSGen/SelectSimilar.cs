using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IBIMSGen
{
    [TransactionAttribute(TransactionMode.Manual)]
    internal class SelectSimilar : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            IList<string> lineStyle = new List<string>();
            IList<ElementId> selection = uidoc.Selection.GetElementIds().ToList().ToList();
            if (selection.Count == 0)
            {
                selection = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select Elements").Select(x => doc.GetElement(x).Id).ToList();
            }
            foreach (ElementId id in selection)
            {
                Element element = doc.GetElement(id);
                if (element != null)
                {
                    if (element is CurveElement)
                    {
                        CurveElement l = (CurveElement)element;
                        lineStyle.Add(l.LineStyle.Name);

                    }
                }
            }
            FilteredElementCollector detailLines = new FilteredElementCollector(doc, doc.ActiveView.Id);
            detailLines.OfCategory(BuiltInCategory.OST_Lines);

            foreach (Element line in detailLines)
            {
                if (line != null)
                {
                    if (line is CurveElement)
                    {
                        CurveElement nl = (CurveElement)line;
                        if (lineStyle.Contains(nl.LineStyle.Name))
                        {
                            selection.Add(line.Id);
                        }
                    }
                }
            }
            Collection<ElementId> newSelection = new Collection<ElementId>(selection);
            uidoc.Selection.SetElementIds(newSelection);

            return Result.Succeeded;
        }
    }
}
