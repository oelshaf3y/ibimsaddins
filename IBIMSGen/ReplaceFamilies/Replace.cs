using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IBIMSGen.ReplaceFamilies
{
    [TransactionAttribute(TransactionMode.Manual)]
    internal class Replace : IExternalCommand
    {
        UIDocument uidoc;
        Document doc;
        Element source, destination;
        List<Element> elems;
        List<XYZ> locations;
        FamilyInstance fI;
        FamilySymbol fs;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            locations = new List<XYZ>();
            ReplaceUI UI = new ReplaceUI();
            UI.ShowDialog();
            if (UI.DialogResult == DialogResult.Cancel) return Result.Failed;
            if (!UI.radioButton3.Checked)
            {

                try
                {
                    td("Select Element to Replace");
                    source = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Pick 1st Element"));
                    td("Select Destination Element");
                    destination = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Pick 2nd Element"));
                    fI = destination as FamilyInstance;
                    fs = fI.Symbol;
                }
                catch (Exception e)
                {
                    return Result.Failed;
                }
            }
            if (UI.radioButton1.Checked)
            {
                elems = new List<Element>();
                elems.Add(source);
                if (!getLocations(elems)) return Result.Failed;
            }
            else if (UI.radioButton2.Checked)
            {
                elems = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategoryId(source.Category.Id).Where(x => x.Name == source.Name).ToList();
                if(!getLocations(elems)) return Result.Failed;
            }
            else
            {
                try
                {

                    td("Select Element to Change");
                    elems = uidoc.Selection.PickObjects(ObjectType.Element, "Pick Elements").Select(x => doc.GetElement(x)).ToList();
                    td("Select Destination Element");
                    destination = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Pick 2nd Element"));
                    fI = destination as FamilyInstance;
                    fs = fI.Symbol;
                }
                catch
                {
                    return Result.Failed;
                }
                if(!getLocations(elems)) return Result.Failed;
            }
            Transaction tr = new Transaction(doc, "Replace");
            tr.Start();

            for (int i = 0; i < locations.Count; i++)
            {
                if (!replaceElement(locations[i], fs, elems[i])) return Result.Failed;
            }
            tr.Commit();
            tr.Dispose();
            //uidoc.Selection.SetElementIds(elems.Select(x=> x.Id).ToList());
            return Result.Succeeded;
        }

        private bool getLocations(List<Element> elems)
        {
            foreach (Element elem in elems)
            {
                try
                {

                    LocationPoint location = elem.Location as LocationPoint;
                    locations.Add(location.Point);
                }
                catch
                {
                    td("Element has no location point\noperation not valid!!".ToUpper());
                    return false;
                }
            }
            return true;
        }

        private bool replaceElement(XYZ location, FamilySymbol fs, Element elem)
        {
            try
            {

                doc.Create.NewFamilyInstance(location, fs, doc.ActiveView);
            }
            catch
            {
                try
                {

                    FamilyInstance famins = elem as FamilyInstance;
                    Element host = famins.Host;
                    doc.Create.NewFamilyInstance(location, fs, host, famins.StructuralType);
                }
                catch
                {
                    try
                    {

                        FamilyInstance famins = elem as FamilyInstance;
                        Reference host = famins.HostFace;
                        doc.Create.NewFamilyInstance(host, location, famins.FacingOrientation, fs);
                    }
                    catch
                    {
                        td("Source element is not face based not element based nor point based. what should i do ?!!!");
                        return false;
                    }

                }
            }
            doc.Delete(elem.Id);
            return true;
        }

        void td(string message)
        {
            TaskDialog.Show("message", message);
        }
    }
}
