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
            List<List<XYZ>> sets = new List<List<XYZ>>();
            foreach (XYZ point in points)
            {
                addToSets(point, sets);
            }

            using (Transaction transaction = new Transaction(doc, "Create cut lines"))
            {
                transaction.Start();
                foreach (List<XYZ> set in sets)
                {
                    XYZ min = set.OrderBy(x => x.DistanceTo(XYZ.Zero)).First();
                    XYZ max = set.OrderByDescending(x => x.DistanceTo(XYZ.Zero)).First();
                    double dx = Math.Abs((max - min).DotProduct(XYZ.BasisX)) + 200 / 304.8;
                    double dy = Math.Abs((max - min).DotProduct(XYZ.BasisY)) + 200 / 304.8;
                    XYZ origin;
                    try
                    {

                        origin = Line.CreateBound(min, max).Evaluate(0.5, true);
                    }
                    catch
                    {
                        origin = min.Add((max - min).GetLength() / 2 * (max - min).Normalize());
                    }
                    FamilyInstance instance = doc.Create.NewFamilyInstance(origin, familySymbol, doc.ActiveView);
                    instance.LookupParameter("Width").Set(Math.Max(dx, dy));
                    instance.LookupParameter("Length").Set(Math.Min(dx, dy));
                    if (dx < dy) instance.Location.Rotate(Line.CreateUnbound(origin, XYZ.BasisZ), Math.PI / 2);

                }
                transaction.Commit();
                transaction.Dispose();
            }

            return Result.Succeeded;
        }

        private void addToSets(XYZ point, List<List<XYZ>> sets)
        {
            bool added = false;
            foreach (List<XYZ> set in sets)
            {
                added = set.Where(x => x.DistanceTo(point) <= 210 / 304.8).Any();
                if (added)
                {
                    set.Add(point);
                    return;
                }

            }
            sets.Add(new List<XYZ> { point });
            return;
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


}
