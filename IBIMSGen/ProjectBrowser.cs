using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IBIMSGen
{
    [Transaction(TransactionMode.Manual)]
    internal class ProjectBrowser : IExternalCommand
    {
        UIDocument uidoc;
        Document doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;

            View activeView = doc.ActiveView as View;
            var f = BrowserOrganization.GetCurrentBrowserOrganizationForViews(doc);
            TaskDialog.Show("Test", f.SortingOrder.ToString());
            return Result.Succeeded;
        }
    }
}
