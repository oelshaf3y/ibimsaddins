using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IBIMSGen
{
    [TransactionAttribute(TransactionMode.Manual)]
    internal class AlignBetween2Pts : IExternalCommand
    {
        UIDocument uidoc;
        Document doc;
        Options options;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            options = new Options();
            options.ComputeReferences = true;
            if (doc.ActiveView is View3D) { TaskDialog.Show("Error", "Active View Can't be 3D!"); return Result.Failed; }
            Element elem = null;
            elem = uidoc.Selection.GetElementIds().Select(x => doc.GetElement(x)).First();
            if (elem == null)
            {
                try
                {
                    elem = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select Elements you'd like to center"));
                }
                catch
                {
                    return Result.Cancelled;
                }
            }
            using (Transaction transaction = new Transaction(doc, "Center Element"))
            {

                transaction.Start();

                Solid solid = getSolid(elem);
                EdgeArray edgeArray = solid.Edges;
                List<XYZ> points = new List<XYZ>();
                foreach (Edge edge in edgeArray)
                {
                    Curve curve = edge.AsCurve();
                    points.Add(curve.GetEndPoint(0));
                    points.Add(curve.GetEndPoint(1));
                }
                XYZ p0 = points.OrderBy(p => p.DistanceTo(XYZ.Zero)).FirstOrDefault();
                XYZ p00 = points.OrderByDescending(p => p.DistanceTo(XYZ.Zero)).FirstOrDefault();
                XYZ centroid = Line.CreateBound(p0, p00).Evaluate(0.5, true);

                Plane plane = Plane.CreateByNormalAndOrigin(uidoc.ActiveView.ViewDirection, uidoc.ActiveView.Origin);

                SketchPlane sp = SketchPlane.Create(doc, plane);

                uidoc.ActiveView.SketchPlane = sp;

                uidoc.ActiveView.ShowActiveWorkPlane();

                XYZ origin = ((LocationPoint)elem.Location).Point;
                //DirectShape.CreateElement(doc,new ElementId(BuiltInCategory.OST_GenericModel)).SetShape(new List<GeometryObject> { Line.CreateBound(XYZ.Zero, origin) });
                XYZ p1, p2, p3, p4;
                try
                {
                    p1 = uidoc.Selection.PickPoint(Autodesk.Revit.UI.Selection.ObjectSnapTypes.Intersections, "Pick 1st corner.");
                    p2 = uidoc.Selection.PickPoint(Autodesk.Revit.UI.Selection.ObjectSnapTypes.Intersections, "Pick 2nd corner.");
                    p3 = new XYZ(p1.X, p1.Y, origin.Z);
                    p4 = new XYZ(p2.X, p2.Y, origin.Z);
                }
                catch
                {
                    return Result.Cancelled;
                }
                XYZ newLocation;
                try
                {
                    newLocation = Line.CreateBound(p3, p4).Evaluate(0.5, true);
                }
                catch
                {
                    newLocation = p3.Add((p4 - p3).GetLength() / 2 * (p4 - p3).Normalize());
                }
                StringBuilder sb = new StringBuilder();

                XYZ delta = (newLocation - centroid);
                XYZ dir = delta.Normalize();
                sb.AppendLine(origin.ToString());
                sb.AppendLine(newLocation.ToString());
                sb.AppendLine(delta.ToString());
                elem.Location.Move(delta);
                sb.AppendLine(((LocationPoint)elem.Location).Point.ToString());
                //TaskDialog.Show("info", sb.ToString());
                transaction.Commit();
                transaction.Dispose();
            }
            //TaskDialog.Show(" s", p.ToString());
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
    }
}
