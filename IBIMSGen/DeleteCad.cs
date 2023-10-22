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
    [TransactionAttribute(TransactionMode.Manual)]
    internal class DeleteCad : IExternalCommand
    {
        UIDocument uidoc;
        Document doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;


            FilteredElementCollector fec = new FilteredElementCollector(doc).OfClass(typeof(CADLinkType));
            if(fec.Count()==0 ) { TaskDialog.Show("Info", "No more DWG Imports In The Project."); }
            Transaction tr = new Transaction(doc, "Delete CAD Imports");
            tr.Start();
            doc.Delete(fec.Select(x => x.Id).ToArray());
            tr.Commit();
            tr.Dispose();
            return Result.Succeeded;
        }
    }
}
