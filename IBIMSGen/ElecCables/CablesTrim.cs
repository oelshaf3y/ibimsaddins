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
        List<Spline> splines, splineSets;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            linesFEC = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_Lines).Where(l => l is CurveElement).Cast<CurveElement>().ToList();
            familySymbol = null;
            familySymbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol))
                ?.Cast<FamilySymbol>()?.Where(x => x != null && x.Name.ToLower().Contains("cut"))?.FirstOrDefault();
            splines = new List<Spline>();
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

            foreach (Element elm in selectedElements)
            {
                if (elm is CurveElement)
                {
                    CurveElement curveElement = (CurveElement)elm;
                    Curve curve = curveElement.GeometryCurve;
                    List<XYZ> intersectionPoints = getIntersections(curve, linesFEC.Select(x => x.GeometryCurve).Where(x => (x is Line)).ToList())
                        ?.OrderBy(x => x.DistanceTo(curve.GetEndPoint(0))).ToList();
                    td(intersectionPoints.Count.ToString());
                    List<List<XYZ>> sets = new List<List<XYZ>>();
                    sets.AddRange(getIntersectionSets(intersectionPoints));
                    foreach (List<XYZ> set in sets)
                    {
                        XYZ pt1st = set.First();
                        XYZ ptLast = set.Last();
                        double length;
                        XYZ origin;
                        if (pt1st.IsAlmostEqualTo(ptLast))
                        {
                            length = 200 / 304.8;
                            origin = pt1st;
                        }
                        else if (pt1st.DistanceTo(ptLast) <= 200)
                        {
                            XYZ dir = (ptLast - pt1st);
                            Line l = Line.CreateBound(pt1st.Add(-100 / 304.8 * dir), ptLast.Add(100 / 304.8 * dir));
                            length = l.Length;
                            origin = l.Evaluate(0.5, true);
                        }
                        else
                        {
                            length = pt1st.DistanceTo(ptLast) + (200 / 304.8);
                            origin = Line.CreateBound(pt1st, ptLast).Evaluate(0.5, true);
                        }
                        double angle = ((Line)curve).Direction.AngleTo(XYZ.BasisY);
                        //FamilyInstance famins = doc.Create.NewFamilyInstance(origin, familySymbol, doc.ActiveView);
                        //famins.LookupParameter("Length").Set(length);
                        //famins.Location.Rotate(Line.CreateUnbound(origin, XYZ.BasisZ), angle);
                        splines.Add(new Spline(curveElement, length, 200, angle, origin));
                    }
                    DetailElementOrderUtils.BringForward(doc, doc.ActiveView, curveElement.Id);
                }
            }

            List<Spline> Splines = splines.OrderBy(x => x.origin.DistanceTo(XYZ.Zero)).ToList();
            Spline prev = Splines.Last();
            for (int i = Splines.Count-2; i >=0; i--)
            {
                Spline spline = Splines[i];
                if (spline.origin.DistanceTo(prev.origin) < 500 / 304.8)
                {
                    //Line l = 
                }
            }

            using (Transaction tr = new Transaction(doc, "Cable Cut"))
            {
                tr.Start();
                familySymbol.Activate();

                tr.Commit();
                tr.Dispose();
            }


            return Result.Succeeded;
        }

        private List<List<XYZ>> getIntersectionSets(List<XYZ> intersectionPoints)
        {
            XYZ prevPoint = XYZ.Zero;
            double distance = 0;
            List<List<XYZ>> sets = new List<List<XYZ>>();
            List<XYZ> set = new List<XYZ>();
            for (int i = 0; i < intersectionPoints.Count; i++)
            {
                if (i != 0) prevPoint = intersectionPoints[i - 1];
                distance = prevPoint.DistanceTo(intersectionPoints[i]);
                if (distance <= 500 / 304.8)
                {
                    set.Add(intersectionPoints[i]);
                }
                else
                {
                    if (set.Count > 0)
                    {
                        sets.Add(set.OrderBy(x => x.DistanceTo(XYZ.Zero)).ToList());
                    }
                    set = new List<XYZ>
                    {
                        intersectionPoints[i]
                    };
                }
            }
            sets.Add(set);
            return sets;
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
        public CurveElement curveElement;
        public double length, width, angle;
        public XYZ origin;
        public bool merged;
        public Spline(CurveElement curveElement, double length, double width, double angle, XYZ origin)
        {
            this.curveElement = curveElement;
            this.length = length;
            this.width = width;
            this.angle = angle;
            this.origin = origin;
            this.merged = false;
        }
    }

}
