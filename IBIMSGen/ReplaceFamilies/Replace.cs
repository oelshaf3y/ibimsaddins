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
using System.Windows.Media.Imaging;

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
        Options options;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            locations = new List<XYZ>();
            options = new Options();
            options.ComputeReferences = true;
            ReplaceUI UI = new ReplaceUI();
            UI.ShowDialog();
            if (UI.DialogResult == DialogResult.Cancel) return Result.Failed;
            if (!UI.radioButton3.Checked)
            {

                try
                {
                    source = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Pick 1st Element"));
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
                if (!getLocations(elems)) return Result.Failed;
            }
            else
            {
                try
                {

                    elems = uidoc.Selection.PickObjects(ObjectType.Element, "Pick Elements").Select(x => doc.GetElement(x)).ToList();
                    destination = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Pick 2nd Element"));
                    fI = destination as FamilyInstance;
                    fs = fI.Symbol;
                }
                catch
                {
                    return Result.Failed;
                }
                if (!getLocations(elems)) return Result.Failed;
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
            FamilyInstance fam;
            var rot = ((LocationPoint)elem.Location).Rotation;
            try
            {

                fam = doc.Create.NewFamilyInstance(location, fs, doc.ActiveView);
            }
            catch
            {
                try
                {
                    FamilyInstance famins = elem as FamilyInstance;
                    Reference host = famins.HostFace;
                    Element el = doc.GetElement(host);
                    if (el is RevitLinkInstance)
                    {
                        RevitLinkInstance rli = (RevitLinkInstance)el;
                        Document linkedDoc = rli.GetLinkDocument();
                        el = linkedDoc.GetElement(host.LinkedElementId);
                        Solid solid = getSolid(el);
                        List<PlanarFace> faces = new List<PlanarFace>();
                        foreach (Face f in solid.Faces)
                        {
                            if (f != null && f is PlanarFace planarFace)
                            {
                                faces.Add(planarFace);
                            }
                        }
                        Face face = faces?.OrderBy(x => x.Origin.DistanceTo(location))?.FirstOrDefault();
                        Reference nHost = face.Reference;
                        if (nHost != null)
                        {

                            fam = doc.Create.NewFamilyInstance(nHost, location, famins.FacingOrientation, fs);
                        }
                        else
                        {
                            try
                            {

                                fam = doc.Create.NewFamilyInstance(face, location, famins.FacingOrientation, fs);
                            }
                            catch
                            {
                                fam = doc.Create.NewFamilyInstance(location, fs, el, famins.StructuralType);

                            }
                        }
                        //host = host.CreateReferenceInLink();
                    }
                    fam = doc.Create.NewFamilyInstance(host, location, famins.FacingOrientation, fs);



                }
                catch (Exception ex)
                {
                    try
                    {
                        FamilyInstance famins = elem as FamilyInstance;
                        Element host = famins.Host;
                        if (host is RevitLinkInstance)
                        {
                            RevitLinkInstance rli = host as RevitLinkInstance;
                            Document linkedDoc = rli.GetLinkDocument();
                            host = linkedDoc.GetElement(host.Id);
                        }

                        fam = doc.Create.NewFamilyInstance(location, fs, host, famins.StructuralType);
                    }
                    catch
                    {
                        td("Source element is not face based not element based nor point based. what should i do ?!!!");
                        return false;
                    }

                }
            }
            fam.Location.Rotate(Line.CreateUnbound(location,XYZ.BasisZ),rot);
            doc.Delete(elem.Id);
            uidoc.Selection.SetElementIds(new List<ElementId>() { fam.Id });
            return true;
        }


        public Solid getSolid(Element elem)
        {
            IList<Solid> solids = new List<Solid>();
            try
            {

                GeometryElement geo = elem.get_Geometry(options);
                if (geo.FirstOrDefault() is Solid)
                {
                    Solid solid = (Solid)geo.FirstOrDefault();
                    return SolidUtils.Clone(solid);
                }
                foreach (GeometryObject geometryObject in geo)
                {
                    if (geometryObject != null)
                    {
                        Solid solid = geometryObject as Solid;
                        if (solid != null && solid.Volume > 0)
                        {
                            solids.Add(solid);

                        }
                    }
                }
            }
            catch
            {
            }
            if (solids.Count == 0)
            {
                try
                {
                    GeometryElement geo = elem.get_Geometry(options);
                    GeometryInstance geoIns = geo.FirstOrDefault() as GeometryInstance;
                    if (geoIns != null)
                    {
                        GeometryElement geoElem = geoIns.GetInstanceGeometry();
                        if (geoElem != null)
                        {
                            foreach (GeometryObject geometryObject in geoElem)
                            {
                                Solid solid = geometryObject as Solid;
                                if (solid != null && solid.Volume > 0)
                                {
                                    solids.Add(solid);
                                }
                            }
                        }
                    }
                }
                catch
                {
                    throw new InvalidOperationException();
                }
            }
            if (solids.Count > 0)
            {
                try
                {

                    return SolidUtils.Clone(solids.OrderByDescending(x => x.Volume).ElementAt(0));
                }
                catch
                {
                    return solids.OrderByDescending(x => x.Volume).ElementAt(0);
                }
            }
            else
            {
                return null;
            }
        }
        void td(string message)
        {
            TaskDialog.Show("message", message);
        }
    }
}
