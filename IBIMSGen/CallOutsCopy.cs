using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IBIMSGen
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CallOutsCopy : IExternalCommand
    {
        UIApplication app;
        Autodesk.Revit.ApplicationServices.Application application;
        UIDocument uidoc;
        Document doc, otherDoc;
        StringBuilder sb = new StringBuilder();
        ICollection<ElementId> callouts, copyIdsViewSpecific, dup;
        ICollection<Element> sections, oViews;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            oViews = new List<Element>();
            app = commandData.Application;
            application = app.Application;
            uidoc = app.ActiveUIDocument;
            doc = uidoc.Document;
            otherDoc = application.Documents.Cast<Document>().Where(d => d.Title != doc.Title).Where(d => d.IsLinked == false).FirstOrDefault();
            if (otherDoc == null)
            {
                TaskDialog.Show("Error", "There must be a 2nd document open.");
                return Result.Failed;
            }
            View activeView = doc.ActiveView;
            ViewSheet activeViewSheet = doc.ActiveView as ViewSheet;
            copyIdsViewSpecific = new Collection<ElementId>();
            callouts = new List<ElementId>();
            sections = new List<Element>();
            dup = new List<ElementId>();

            using (Transaction tx = new Transaction(otherDoc))
            {
                tx.Start("copy callouts");
                if (activeViewSheet == null)
                {
                    td("active view must be a sheet!");
                    return Result.Failed;
                }
                else
                {
                    FilteredElementCollector otherViews = new FilteredElementCollector(otherDoc).OfCategory(BuiltInCategory.OST_Views);
                    //otherViews.WhereElementIsNotElementType();
                    otherViews.ToElements().ToList();
                    oViews = otherViews.ToElements();
                    FilteredElementCollector vports = new FilteredElementCollector(doc).OwnedByView(doc.ActiveView.Id);
                    foreach (Element e in vports)
                    {
                        if (e is Viewport)
                        {
                            Viewport vp = e as Viewport;
                            View v = doc.GetElement(vp.ViewId) as View;
                            if (v != null && v.GetType() != typeof(ViewSheet) && v.GetType() != typeof(ViewSchedule)
                    && v.ViewType != ViewType.ProjectBrowser && v.ViewType != ViewType.SystemBrowser && v.ViewType != ViewType.Legend)
                            {
                                //sb.AppendLine(v.Name);
                                copyCallouts(v);

                            }
                        }
                    }
                }
                //td(sb.ToString());
                tx.Commit();
                tx.Dispose();
                td("Callouts copied successfully");
            }
            return Result.Succeeded;
        }
        // message
        public void td(string message)
        {
            TaskDialog.Show("message", message);
        }

        // get section line 
        public Line GetSectionLine(Element section, View view)
        {
            try
            {
                const double correction = 21.130014403 / 304.8;

                Category cat = section.Category;

                if (null == cat)
                {
                    throw new ArgumentException(
                      "Section has null category");
                }

                if (BuiltInCategory.OST_Viewers
                  != (BuiltInCategory)(cat.Id.IntegerValue))
                {
                    throw new ArgumentException(
                      "Expected section with OST_Viewers category");
                }

                FilteredElementCollector views
                  = new FilteredElementCollector(doc)
                    .OfClass(typeof(View));

                View viewFromSection = null;

                foreach (View v in views)
                {
                    if (section.Name == v.Name
                      && section.GetTypeId() == v.GetTypeId())
                    {
                        viewFromSection = v;
                        break;
                    }
                }
                if (viewFromSection == null) return null;

                ViewFamilyType vType = doc.GetElement(
                  section.GetTypeId()) as ViewFamilyType;

                BoundingBoxXYZ bb1 = null;

                using (Transaction st1 = new Transaction(doc))
                {
                    st1.Start("getLine");
                    Parameter par = vType.get_Parameter(
                      BuiltInParameter.SECTION_TAG);

                    par.Set(ElementId.InvalidElementId);

                    par = vType.get_Parameter(
                      BuiltInParameter.VIEWER_REFERENCE_LABEL_TEXT);

                    par.Set(string.Empty);
                    view.Scale = 1;
                    doc.Regenerate();
                    bb1 = section.get_BoundingBox(view);
                    st1.RollBack();
                }

                BoundingBoxXYZ bb = section.get_BoundingBox(view);
                XYZ pt1 = bb.Min;
                XYZ pt2 = bb.Max;
                if (bb1 != null)
                {
                    pt1 = bb1.Min;
                    pt2 = bb1.Max;
                }
                XYZ Origin = viewFromSection.Origin;
                XYZ ViewBasisX = viewFromSection.RightDirection;
                XYZ ViewBasisY = viewFromSection.ViewDirection;
                XYZ ViewBasisZ = viewFromSection.UpDirection;
                if (ViewBasisX.X < 0 ^ ViewBasisX.Y < 0)
                {
                    double d = pt1.Y;
                    pt1 = new XYZ(pt1.X, pt2.Y, pt1.Z);
                    pt2 = new XYZ(pt2.X, d, pt2.Z);
                }
                XYZ ToPlane1, ToPlane2, correctionVector;


                if (view.ViewType == ViewType.Elevation)
                {
                    ToPlane1 = pt1.Add(ViewBasisZ.Multiply(
                  ViewBasisZ.DotProduct(Origin.Subtract(pt1))));

                    ToPlane2 = pt2.Subtract(ViewBasisZ.Multiply(
                      ViewBasisZ.DotProduct(pt2.Subtract(Origin))));

                    correctionVector = ToPlane2.Subtract(ToPlane1)
                      .Normalize().Multiply(correction);
                    //pt1 = new XYZ(pt1.X, pt1.Z, pt1.Y);
                    //pt2 = new XYZ(pt2.X, pt2.Z, pt2.Y);

                }
                else
                {
                    ToPlane1 = pt1.Add(ViewBasisY.Multiply(
                  ViewBasisY.DotProduct(Origin.Subtract(pt1))));

                    ToPlane2 = pt2.Subtract(ViewBasisY.Multiply(
                      ViewBasisY.DotProduct(pt2.Subtract(Origin))));

                    correctionVector = ToPlane2.Subtract(ToPlane1)
                      .Normalize().Multiply(correction);
                }

                XYZ endPoint0 = ToPlane1.Add(correctionVector);
                XYZ endPoint1 = ToPlane2.Subtract(correctionVector);
                XYZ p1, p2;
                if (Math.Abs(endPoint0.X - endPoint1.X) < .1)
                {
                    p1 = new XYZ(endPoint0.X, endPoint0.Y, endPoint0.Z);
                    p2 = new XYZ(endPoint0.X, endPoint1.Y, endPoint1.Z);
                }
                else if (Math.Abs(endPoint0.Y - endPoint1.Y) < .1)
                {
                    p1 = new XYZ(endPoint0.X, endPoint0.Y, endPoint0.Z);
                    p2 = new XYZ(endPoint1.X, endPoint0.Y, endPoint1.Z);
                }
                else
                {
                    p1 = new XYZ(endPoint0.X, endPoint0.Y, endPoint0.Z);
                    p2 = new XYZ(endPoint1.X, endPoint1.Y, endPoint1.Z);
                }
                return Line.CreateBound(p1, p2);
            }
            catch
            {
                return null;
            }
        }

        public void copyCallouts(View v)
        {
            bool found = false;
            callouts = v.GetReferenceCallouts();
            FilteredElementCollector sc = new FilteredElementCollector(doc).OwnedByView(v.Id);
            sc.OfCategory(BuiltInCategory.OST_Viewers).ToElements().ToList();

            if (callouts.Count() > 0)
            {

                foreach (Element section in sc)
                {
                    if (!callouts.Contains(section.Id))
                    {
                        if (section != null)
                        {
                            sections.Add(section);
                        }
                    }
                }
                foreach (Element ee in oViews)
                {
                    if (ee is View)
                    {
                        View view = ee as View;
                        if (view.Name == v.Name)
                        {
                            found = true;
                            foreach (ElementId id in callouts)
                            {
                                Element el2 = doc.GetElement(id);
                                View calloutView = el2 as View;
                                BoundingBoxXYZ bx = el2.get_BoundingBox(v);
                                XYZ p1 = new XYZ(bx.Min.X, bx.Min.Y, bx.Min.Z);
                                XYZ p2 = new XYZ(bx.Max.X, bx.Max.Y, bx.Max.Z);
                                foreach (Element el in oViews)
                                {
                                    if (el.Name.Trim().Remove(el.Name.Length - 1).Trim() == el2.Name.Trim().Remove(el2.Name.Length - 1).Trim())
                                    {
                                        //sb.Append(el.Name);
                                        ViewSection.CreateReferenceCallout(otherDoc, view.Id, el.Id, p1, p2);
                                    }
                                }
                            }
                            foreach (Element sec in sections)
                            {
                                Line l = GetSectionLine(sec, v);
                                if (l != null)
                                {
                                    foreach (Element el in oViews)
                                    {
                                        if (el.Name.Trim().Remove(el.Name.Length - 1).Trim() == sec.Name.Trim().Remove(sec.Name.Length - 1).Trim())
                                        {
                                            //sb.Append(el.Name);
                                            ViewSection.CreateReferenceSection(otherDoc, view.Id, el.Id, l.GetEndPoint(1), l.GetEndPoint(0));
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                sections = sc.ToElements();
                td("no callouts found in the view: " + v.Name);
                if (sections.Count() > 0)
                {
                    foreach (Element ee in oViews)
                    {
                        if (ee is View)
                        {
                            View view = ee as View;
                            if (view.Name == v.Name)
                            {
                                foreach (Element sec in sections)
                                {
                                    Line l = GetSectionLine(sec, v);
                                    if (l != null)
                                    {
                                        foreach (Element el in oViews)
                                        {
                                            if (el.Name.Trim().Remove(el.Name.Length - 1).Trim() == sec.Name.Trim().Remove(sec.Name.Length - 1).Trim())
                                            {
                                                sb.Append(el.Name);
                                                ViewSection.CreateReferenceSection(otherDoc, view.Id, el.Id, l.GetEndPoint(1), l.GetEndPoint(0));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                return;
            }
            if (!found)
            {
                td("the view: " + v.Name + " is not found in the project: " + otherDoc.PathName.Split('\\').Last());
            }
        }
    }
}
