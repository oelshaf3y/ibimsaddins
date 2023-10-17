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

namespace IBIMSGen.ElecCables
{
    [TransactionAttribute(TransactionMode.Manual)]
    internal class CablesTrim : IExternalCommand
    {
        UIDocument uidoc;
        Document doc;
        FamilySymbol familySymbol;
        List<Element> selectedElements;
        List<CurveElement> linesFEC;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            linesFEC = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_Lines).Where(l => l is CurveElement).Cast<CurveElement>().ToList();
            familySymbol = null;
            familySymbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol))
                ?.Cast<FamilySymbol>()?.Where(x => x != null && x.Name.ToLower().Contains("cut"))?.FirstOrDefault();
            if (familySymbol == null)
            {
                td("Loading Family");
                using (Transaction t = new Transaction(doc))
                {

                    t.Start("loading Spline");
                    try
                    {
                        doc.LoadFamily(@"H:\01-Eng\8-Others\IBIMS Addins\Omar\Families\Electrical\CableCut.rfa");
                        doc.Regenerate();
                        familySymbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol))
                            .Cast<FamilySymbol>().Where(x => x != null && x.Name.ToLower().Contains("cut"))?.FirstOrDefault();
                        t.Commit();
                        t.Dispose();
                    }
                    catch (Exception ex)
                    {
                        td(ex.Message);
                        t.RollBack();
                    }
                }
            }

            try
            {
                selectedElements = uidoc.Selection.PickObjects(ObjectType.Element, new LineSelectionFilter(), "Select Electrical Detail Lines")
                    .Select(x => doc.GetElement(x)).ToList();
            }
            catch
            {
                return Result.Cancelled;
            }
            double z = 0;
            List<XYZ> points = new List<XYZ>();
            foreach (Element elm in selectedElements)
            {
                if (elm is CurveElement)
                {
                    CurveElement curveElement = (CurveElement)elm;
                    Curve curve = curveElement.GeometryCurve;
                    points.AddRange(getIntersections(curve, linesFEC.Select(x => x.GeometryCurve).Where(x => (x is Line)).ToList())
                        ?.OrderBy(x => x.DistanceTo(curve.GetEndPoint(0))).ToArray());
                    if (z == 0) z = points[0].Z;
                }
            }
            double minx, miny, maxx, maxy;

            minx = points.OrderBy(x => x.X).First().X;
            miny = points.OrderBy(x => x.Y).First().Y;
            maxx = points.OrderByDescending(x => x.X).First().X;
            maxy = points.OrderByDescending(x => x.Y).First().Y;

            //td(points.Count.ToString());
            QuadTree qtree = new QuadTree(new Rectangle(Convert.ToInt32(minx), Convert.ToInt32(maxy), Convert.ToInt32(maxx), Convert.ToInt32(miny)));
            foreach (XYZ point in points.Distinct().ToList())
            {
                qtree.Insert(point);
            }
            List<List<XYZ>> sets = new List<List<XYZ>>();
            List<XYZ> nearest = new List<XYZ>();
            int range = Convert.ToInt32(3000 / 304.8);
            foreach (XYZ point in points)
            {
                if (CollectedPoint(sets, point)) continue;
                if (nearest.Where(x => x.IsAlmostEqualTo(point)).Any()) continue;
                QuadTree rect = new QuadTree(new Rectangle(Convert.ToInt32(point.X) - range, Convert.ToInt32(point.Y) + range, Convert.ToInt32(point.X) + range, Convert.ToInt32(point.Y) - range));
                List<XYZ> query = qtree.queryRange(rect);
                if (query.Count > 0)
                {
                    nearest.AddRange(query.ToArray());
                }
                else
                {
                    if (nearest.Count > 0) sets.Add(nearest.Distinct().ToList());
                    nearest = new List<XYZ>();
                }
            }
            //td(sets.Count.ToString());
            Transaction tr = new Transaction(doc);
            tr.Start("a7a");
            familySymbol.Activate();
            foreach (List<XYZ> set in sets)
            {
                XYZ min = set.OrderBy(x => x.DistanceTo(XYZ.Zero)).First()
                    , max = set.OrderByDescending(x => x.DistanceTo(XYZ.Zero)).First(),
                    origin = Line.CreateBound(min, max).Evaluate(0.5, true);
                FamilyInstance fam = doc.Create.NewFamilyInstance(origin, familySymbol, doc.ActiveView);
                double dx = (max - min).DotProduct(XYZ.BasisX);
                double dy = (max - min).DotProduct(XYZ.BasisY);
                fam.LookupParameter("Width").Set(Math.Max(dx, dy) + 400 / 304.8);
                fam.LookupParameter("Length").Set((Math.Min(dx, dy)) + 400 / 304.8);
                if (dy > dx) fam.Location.Rotate(Line.CreateUnbound(origin, XYZ.BasisZ), Math.PI / 2);
            }

            tr.Commit();
            tr.Dispose();
            return Result.Succeeded;
        }

        private bool CollectedPoint(List<List<XYZ>> sets, XYZ point)
        {
            foreach (List<XYZ> set in sets)
            {
                if (set.Where(x => x.IsAlmostEqualTo(point)).Any()) return true;
            }
            return false;
        }

        private List<XYZ> getIntersections(Curve curve, List<Curve> curves)
        {
            List<XYZ> intersectionPoints = new List<XYZ>();
            foreach (Curve otherCurve in curves)
            {
                IntersectionResultArray ira = null;
                curve.Intersect(otherCurve, out ira);
                if (ira != null && !ira.IsEmpty)
                {
                    XYZ point = ira.get_Item(0).XYZPoint;
                    intersectionPoints.Add(point);
                }
            }
            return intersectionPoints;
        }

        void td(string message)
        {
            TaskDialog.Show("Message", message);
        }
    }

    public class Spline
    {
        double Width;
        double Length;
        double Angle;
        XYZ Origin;
        public Spline(double width, double length, double angle, XYZ origin)
        {
            Width = width;
            Length = length;
            Angle = angle;
            Origin = origin;
        }
    }

}
