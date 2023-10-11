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
    internal class BlockWork : IExternalCommand
    {
        UIDocument uidoc;
        Document doc;
        Options options;
        StringBuilder sb;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            options = new Options();
            options.ComputeReferences = true;
            FilteredElementCollector walls = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType();
            FilteredElementCollector dims = new FilteredElementCollector(doc, doc.ActiveView.Id).OfClass(typeof(Dimension)).WhereElementIsNotElementType();
            FilteredElementCollector openings=new FilteredElementCollector(doc,doc.ActiveView.Id).OfClass(typeof(Opening)).WhereElementIsNotElementType();
            uidoc.Selection.SetElementIds(openings.Select(x=> x.Id).ToList());
            return Result.Succeeded;
            Dimension d = dims.ToElements()?.FirstOrDefault(e => e is Dimension && !(e is SpotDimension)) as Dimension;
            if (d == null) return Result.Cancelled;
            List<Wall> collection = new List<Wall>();
            double elevation = (d.Curve as Line).Origin.Z;

            sb = new StringBuilder();
            foreach (Element w in walls)
            {
                if (w is Wall wall)
                {
                    if (wall.Width < 100 / 304.8) continue;
                    Solid solid = getSolid(wall);
                    if (solid != null)
                    {
                        LocationCurve location = wall.Location as LocationCurve;
                        Line curve = location.Curve as Line;
                        if (curve == null) continue;
                        XYZ dir = curve.Direction;
                        XYZ point = new XYZ(curve.GetEndPoint(0).X, curve.GetEndPoint(0).Y, elevation);
                        Curve ray = Line.CreateUnbound(point, dir) as Curve;
                        IntersectionResultArray irr = null;
                        foreach (Face face in solid.Faces)
                        {
                            if (face.Intersect(ray, out irr) != SetComparisonResult.Disjoint)
                            {
                                if (irr != null)
                                {

                                    if (!irr.IsEmpty)
                                    {
                                        PlanarFace pf = face as PlanarFace;
                                        if (pf.Origin == curve.GetEndPoint(0) || pf.Origin == curve.GetEndPoint(1)) continue;
                                        collection.Add(wall);
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            uidoc.Selection.SetElementIds(collection.Select(x=>x.Id).ToList()); 
            return Result.Succeeded;
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
            TaskDialog.Show("Message", message);
        }
    }
}
