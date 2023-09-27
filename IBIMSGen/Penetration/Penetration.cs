using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static IBIMSGen.SelectionExtentions;

namespace IBIMSGen
{
    [Transaction(TransactionMode.Manual)]
    public partial class Penetration : IExternalCommand
    {
        UIApplication app;
        UIDocument uidoc;
        Document doc, linkedDoc;
        FilteredElementCollector famSymbolFEC;
        IList<FamilySymbol> famSymbols;
        IList<Element> mechElements, strElement;
        FamilySymbol SelectedFamSymbol;
        StringBuilder sb;
        Options options;
        IList<Reference> references;
        IList<RevitLinkInstance> rli;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            app = commandData.Application;
            uidoc = app.ActiveUIDocument;
            doc = uidoc.Document;
            options = new Options();
            options.ComputeReferences = true;
            sb = new StringBuilder();
            famSymbolFEC = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol));
            famSymbols = famSymbolFEC.WhereElementIsElementType().Select(x => x as FamilySymbol).ToList();
            SelectedFamSymbol = famSymbols.Where(x => x.Name.ToLower().Equals("ct")).FirstOrDefault();
            td("Select Penetrating Elements\nDucts, Pipes, Cable Trays or Conduits");
            mechElements = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, new mechSelectionFilter(), "Select Mechanical Elements").Select(x => doc.GetElement(x)).ToList();
            bool strLink = true;
            td("Select Structural Elements which will be penetrated.\nFloors, walls, beams or columns");
            if (strLink)
            {
                references = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.LinkedElement, new strSelectionFilter(), "Select Structural Elements");
                rli = references.Select(x => (doc.GetElement(x.ElementId) as RevitLinkInstance)).ToList();
                strElement = references.Select(x => ((RevitLinkInstance)doc.GetElement(x.ElementId)).GetLinkDocument().GetElement(x.LinkedElementId)).ToList();
            }
            else
            {
                strElement = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, new strSelectionFilter(), "Select Structural Elements").Select(x => doc.GetElement(x)).ToList();
            }
            Transaction tr = new Transaction(doc, "Penetrate");
            tr.Start();
            foreach (Element elem in mechElements)
            {
                LocationCurve locationCurve = elem.Location as LocationCurve;
                foreach (Element penetrated in strElement)
                {
                    bool isCircular = false;
                    double depth = 0, height = 0, width = 0, insulationThickness = 0;
                    Solid solid = getSolid(penetrated);
                    if (solid == null) continue;
                    SolidCurveIntersectionOptions sci = new SolidCurveIntersectionOptions();
                    sci.ResultType = SolidCurveIntersectionMode.CurveSegmentsInside;
                    try
                    {
                        SolidCurveIntersection solidCurveIntersection = solid.IntersectWithCurve(locationCurve.Curve, sci);
                        if (solidCurveIntersection.SegmentCount > 0)
                        {
                            Curve intersectionCurve = solidCurveIntersection.GetCurveSegment(0);
                            depth = intersectionCurve.Length;
                            try
                            {
                                width = elem.LookupParameter("Diameter").AsDouble();
                                isCircular = true;
                                height = width;
                            }
                            catch
                            {
                                isCircular = false;
                                width = elem.LookupParameter("Width").AsDouble();
                                height = elem.LookupParameter("Height").AsDouble();
                            }
                            try
                            {
                                insulationThickness = elem.LookupParameter("Insulation Thickness").AsDouble();
                            }
                            catch { }
                            width += (2 * insulationThickness) + (50 / 304.8);
                            height += (2 * insulationThickness) + (50 / 304.8);
                            Reference refFace = null;
                            foreach (Face face in solid.Faces)
                            {
                                IntersectionResultArray ira = new IntersectionResultArray();
                                face.Intersect(locationCurve.Curve, out ira);

                                if (ira != null)
                                {
                                    if (!ira.IsEmpty)
                                    {
                                        if (references != null)
                                        {
                                            refFace = face.Reference.CreateLinkReference(rli.ElementAt(strElement.IndexOf(penetrated)));

                                        }
                                        else
                                        {
                                            refFace = face.Reference;

                                        }
                                    }
                                }
                            }
                            Line line = intersectionCurve as Line;
                            XYZ dir = line.Direction;
                            XYZ locationPt;
                            if (dir.X * dir.Y * dir.Z > 0) { locationPt = intersectionCurve.GetEndPoint(0); } else { locationPt = intersectionCurve.GetEndPoint(1); }
                            //locationPt = intersectionCurve.GetEndPoint(1);
                            SelectedFamSymbol.Activate();
                            FamilyInstance familyInstance = null;
                            try
                            {

                                familyInstance = doc.Create.NewFamilyInstance(refFace, locationPt, dir, SelectedFamSymbol);
                            }
                            catch (Exception ex)
                            {
                                td(ex.Message);
                            }

                            if (familyInstance != null)
                            {
                                td(SelectedFamSymbol.Name);
                                if (isCircular)
                                {
                                    familyInstance.LookupParameter("Diameter").Set(width);
                                }
                                else
                                {
                                    familyInstance.LookupParameter("b").Set(width);
                                    familyInstance.LookupParameter("h").Set(height);
                                    //familyInstance.Location.Rotate(line, Math.PI / 2);
                                }
                                familyInstance.LookupParameter("Depth").Set(depth + (50 / 304.8));
                                uidoc.Selection.SetElementIds(new List<ElementId> { SelectedFamSymbol.Id });
                                tr.Commit();
                                tr.Dispose();
                                return Result.Succeeded;
                            }


                        }
                    }
                    catch (Exception ex)
                    {
                        td(ex.Message);
                    }
                }
            }
            tr.Commit();
            tr.Dispose();

            return Result.Succeeded;
        }



        Solid getSolid(Element elem)
        {
            List<Solid> solids = new List<Solid>();
            GeometryElement geo = elem.get_Geometry(options);
            foreach (GeometryObject obj in geo)
            {
                if (obj is Solid)
                {
                    Solid solid = (Solid)obj;
                    if (solid != null && solid.Volume != 0)
                    {
                        solids.Add(solid);
                    }
                }
            }
            if (solids.Count == 0)
            {
                GeometryInstance inst = geo.First() as GeometryInstance;
                foreach (GeometryObject obj in inst.GetInstanceGeometry())
                {
                    if (obj is Solid)
                    {
                        Solid solid = (Solid)obj;
                        if (solid != null && solid.Volume != 0)
                        {
                            solids.Add(solid);
                        }
                    }
                }
            }
            return solids.OrderByDescending(x => x.Volume).FirstOrDefault();
        }

        void td(string message)
        {
            TaskDialog.Show("Message", message);
        }
    }

}
