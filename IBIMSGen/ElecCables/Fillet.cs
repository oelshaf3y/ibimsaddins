using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace IBIMSGen.ElecCables
{
    [TransactionAttribute(TransactionMode.Manual)]
    internal class Fillet : IExternalCommand
    {
        UIDocument uidoc;
        Document doc;
        List<Element> selectedElements;
        string lineStyle;
        List<CurveElement> lines;
        double R;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            R = 2 *doc.ActiveView.Scale / 304.8;
            //Transaction t = new Transaction(doc, "tr");
            //t.Start("direct shape");
            //DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel)).SetShape(
            //    new List<GeometryObject> { NurbSpline.CreateCurve(
            //        new List<XYZ>
            //        {
            //            //2 4 1 3 
            //            new XYZ(26.1349027046458,38.7422687557032,0),
            //            new XYZ(24.7722236317828,36.7318481581823,0),
            //            new XYZ(27.0452561876474,35.419512200177,0),
            //            new XYZ(25.5726351525734,33.2468893719458,0)


            //        }
            //        ,new List<double> { 1,1,1,1})}
            //    );
            //t.Commit();
            //t.Dispose();
            //return Result.Succeeded;

            selectedElements = uidoc.Selection.GetElementIds().Select(x => doc.GetElement(x)).ToList();
            if (selectedElements.Count == 0)
            {
                try
                {
                    selectedElements = uidoc.Selection.PickObjects(ObjectType.Element, new LineSelectionFilter(), "Pick Lines")
                        .Select(x => doc.GetElement(x)).ToList();
                }
                catch
                {
                    return Result.Cancelled;
                }
            }

            using (Transaction tr = new Transaction(doc))
            {
                tr.Start("Fillet");
                for (int i = 0; i < selectedElements.Count; i++)
                {
                    Element element = selectedElements.ElementAt(i);
                    CurveElement selectedElement = element as CurveElement;
                    lineStyle = selectedElement.LineStyle.Name;
                    Curve selectedCurve = selectedElement.GeometryCurve;
                    if (selectedCurve == null) continue;
                    lines = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_Lines)
                        .Where(l => l is CurveElement).Cast<CurveElement>()
                        .Where(l => l.LineStyle.Name == lineStyle).ToList();
                    foreach (CurveElement line in lines)
                    {
                        Curve curve = line.GeometryCurve;
                        if (curve != null)
                        {
                            IntersectionResultArray irr = null;
                            if (!(selectedCurve.Intersect(curve, out irr) == SetComparisonResult.Disjoint))
                            {
                                if (irr == null || irr.IsEmpty) continue;
                                XYZ intersectionPt = irr.get_Item(0).XYZPoint;
                                if (intersectionPt == null) continue;
                                if (
                                    intersectionPt.IsAlmostEqualTo(curve.GetEndPoint(0)) ||
                                    intersectionPt.IsAlmostEqualTo(curve.GetEndPoint(1)) ||
                                    intersectionPt.IsAlmostEqualTo(selectedCurve.GetEndPoint(0)) ||
                                    intersectionPt.IsAlmostEqualTo(selectedCurve.GetEndPoint(1))
                                    )
                                {
                                    Curve firstCurve = fixOrientation(selectedCurve, intersectionPt);
                                    Curve otherCurve = fixOrientation(curve, intersectionPt);
                                    double theta = ((Line)curve).Direction.AngleTo(((Line)selectedCurve).Direction) * 0.5;
                                    XYZ curveDir = ((Line)otherCurve).Direction.Normalize();
                                    XYZ selectedCurveDir = ((Line)firstCurve).Direction.Normalize();
                                    XYZ p1, p2;
                                    p1 = intersectionPt.Add(-R * Math.Tan(theta) * curveDir);
                                    p2 = intersectionPt.Add(-R * Math.Tan(theta) * selectedCurveDir);
                                    XYZ perpendicular = getPerpendicularDir(otherCurve, firstCurve, intersectionPt);
                                    XYZ centerPt = p1.Add(-R * perpendicular);
                                    XYZ ellipseDir1 = Line.CreateBound(p1, centerPt).Direction.Normalize();
                                    XYZ ellipseDir2 = Line.CreateBound(p2, centerPt).Direction.Normalize();
                                    double a1, a2;
                                    a1 = Math.Min(Math.Atan(ellipseDir1.Y / ellipseDir1.X), Math.Atan(ellipseDir2.Y / ellipseDir2.X));
                                    a2 = Math.Max(Math.Atan(ellipseDir1.Y / ellipseDir1.X), Math.Atan(ellipseDir2.Y / ellipseDir2.X));
                                    Curve ellipse = Ellipse.CreateCurve(centerPt, R, R, perpendicular, curveDir, 0, 2 * theta);
                                    DetailCurve arc = doc.Create.NewDetailCurve(doc.ActiveView, ellipse);
                                    arc.LineStyle = line.LineStyle;
                                    DetailCurve c2 = doc.Create.NewDetailCurve(doc.ActiveView, Line.CreateBound(p1, otherCurve.GetEndPoint(0)));
                                    c2.LineStyle = line.LineStyle;
                                    DetailCurve elem2 = doc.Create.NewDetailCurve(doc.ActiveView, Line.CreateBound(p2, firstCurve.GetEndPoint(0)));
                                    elem2.LineStyle = line.LineStyle;
                                    ElementId id1 = line.Id, id2 = element.Id;
                                    selectedElements[i] = doc.GetElement(elem2.Id);
                                    doc.Delete(id1);
                                    doc.Delete(id2);
                                    break;
                                }
                            }
                        }
                    }
                }
                tr.Commit();
                tr.Dispose();

            }



            return Result.Succeeded;
        }

        private XYZ getPerpendicularDir(Curve curve, Curve otherCurve, XYZ intersectionPt)
        {
            XYZ dir = ((Line)curve).Direction.Normalize();
            XYZ perp1 = new XYZ(-dir.Y, dir.X, dir.Z);
            XYZ perp2 = new XYZ(dir.Y, -dir.X, dir.Z);
            XYZ pt1 = intersectionPt.Add(3 * perp1);
            XYZ pt2 = intersectionPt.Add(3 * perp2);
            XYZ curvePt = otherCurve.GetEndPoint(0);
            if (curvePt.DistanceTo(pt1) < curvePt.DistanceTo(pt2))
            {
                return perp2;
            }
            else
            {
                return perp1;
            }
        }

        private Curve fixOrientation(Curve curve, XYZ intersectionPt)
        {
            List<XYZ> pts = new List<XYZ>() { curve.GetEndPoint(0), curve.GetEndPoint(1) }.OrderByDescending(x => x.DistanceTo(intersectionPt)).ToList();
            return Line.CreateBound(pts[0], pts[1]) as Curve;
        }
    }

    public class LineSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem.Category.Id == new ElementId(BuiltInCategory.OST_Lines) && elem is CurveElement;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}
