using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.LinkLabel;

namespace IBIMSGen.ElecCables
{
    [Transaction(TransactionMode.Manual)]
    internal class CableTraysElevations : IExternalCommand
    {
        UIDocument uidoc;
        Document doc;
        List<Element> floors;
        Options options;
        ProjectPosition projectPosition;
        StringBuilder sb = new StringBuilder();
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            options = new Options();
            options.ComputeReferences = true;
            List<RevitLinkInstance> linkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();
            floors = new List<Element>();
            List<Element> trays;
            bool selection = uidoc.Selection.GetElementIds().Count > 0;
            if (selection)
            {
                trays = uidoc.Selection.GetElementIds().Select(x => doc.GetElement(x)).Where(x=>x.Category.Name.Equals("Cable Trays")).ToList();
                sb.AppendLine("Cable trays offset form slab are:\n");
            }
            else
            {
                trays = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType().ToList();
            }

            Element nearest = null;
            XYZ nearestPt = XYZ.Zero;
            XYZ posXYZ = XYZ.Zero;
            foreach (RevitLinkInstance rli in linkInstances)
            {
                Document link = rli.GetLinkDocument();
                floors.AddRange(new FilteredElementCollector(link).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToList());
            }
            using (Transaction tr = new Transaction(doc, "Trays Elevations"))
            {
                tr.Start();
                int c = 0;
                if (trays.Count == 0) { td("No trays found!"); return Result.Cancelled; }
                foreach (Element elem in trays)
                {
                    Line curve = ((LocationCurve)elem.Location).Curve as Line;
                    double min = double.MaxValue;
                    XYZ midPoint = curve.Evaluate(0.5, true);
                    if (Math.Round(curve.Direction.Z) != 0) continue;
                    c++;
                    foreach (Element floor in floors)
                    {
                        if (floor is Floor)
                        {
                            Solid solid = getSolid(floor);
                            Face botFace = solid.Faces.get_Item(0);
                            botFace.Intersect(Line.CreateBound(midPoint.Add(-500 * XYZ.BasisZ), midPoint.Add(500 * XYZ.BasisZ)), out IntersectionResultArray ira);
                            if (ira != null && !ira.IsEmpty)
                            {
                                XYZ intersectionPt = ira.get_Item(0).XYZPoint;
                                if (Math.Abs(intersectionPt.DistanceTo(midPoint)) < min)
                                {
                                    min = intersectionPt.DistanceTo(midPoint);
                                    nearest = floor;
                                    nearestPt = intersectionPt;
                                }

                            }
                        }
                    }

                    double height = elem.LookupParameter("Height").AsDouble();
                    string heightOffset = Math.Round((nearestPt.Z - midPoint.Z - height / 2) * 304.8).ToString();
                    //DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel)).SetShape(new List<GeometryObject> { Line.CreateBound(midPoint, nearestPt) });
                    if (selection) sb.AppendLine(heightOffset + " mm ");
                    elem.LookupParameter("Comments").Set(heightOffset + " mm ");
                }
                if (selection) { sb.AppendLine(" and"); } else { sb.AppendLine("Cable Trays heights offset of this active view"); }

                sb.Append(" have been stored in comments.");
                td(sb.ToString());
                tr.Commit();
                tr.Dispose();
            }
            return Result.Succeeded;
        }

        public void td(string message)
        {
            TaskDialog.Show("info", message);

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
    }
}
