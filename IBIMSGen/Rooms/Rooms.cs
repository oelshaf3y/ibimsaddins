using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IBIMSGen.Rooms
{
    [TransactionAttribute(TransactionMode.Manual)]
    internal class Rooms : IExternalCommand
    {
        UIDocument uidoc;
        Document doc;
        List<Room> rooms;
        Options options;
        StringBuilder sb;
        RoomsUI UI;
        Element viewForSection, viewForDetials;
        ElementId titleBlockId;
        FilteredElementCollector views, titleBlocks, vft;
        List<ElementId> viewTempsIds;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            options = new Options();
            options.ComputeReferences = true;
            viewForSection = null;
            viewForDetials = null;
            views = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views);
            vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType));
            viewTempsIds = views.Cast<View>().Where(x => x.IsTemplate).OrderBy(x => x.Name).Select(x => x.Id).Distinct().ToList();
            titleBlocks = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks);
            titleBlockId = titleBlocks.FirstOrDefault().Id;
            foreach (Element view in vft)
            {
                ViewFamilyType vf = view as ViewFamilyType;
                if (vf.ViewFamily == ViewFamily.Section)
                {
                    viewForSection = view;
                }
                if (vf.ViewFamily == ViewFamily.Detail)
                {
                    viewForDetials = view;
                }
            }
            UI = new RoomsUI(doc, viewTempsIds);
            UI.ShowDialog();
            if (UI.DialogResult == System.Windows.Forms.DialogResult.Cancel)
            {
                return Result.Failed;
            }
            try
            {
                rooms = uidoc.Selection.PickObjects(ObjectType.Element, new RoomSelectionFilter(), "select rooms").Select(x => doc.GetElement(x) as Room).ToList();
            }
            catch
            {
                return Result.Failed;
            }

            sb = new StringBuilder();
            double min = double.MaxValue;
            double max = double.MinValue;
            List<ElementId> ids = new List<ElementId>();
            Transaction tr = new Transaction(doc);
            tr.Start("wall section");
            foreach (Room room in rooms)
            {
                List<View> viewPorts = new List<View>();
                ViewSheet sheet = ViewSheet.Create(doc, titleBlockId);
                try
                {

                    if (UI.sheetNumber != null) sheet.SheetNumber = UI.sheetNumber;
                }
                catch { }
                if (UI.sheetName != null)
                {
                    sheet.Name = UI.sheetName;
                }
                else
                {
                    sheet.Name = room.Name;
                }
                Solid solid = getSolid(room);
                double minx = double.MaxValue, miny = double.MaxValue, maxx = double.MinValue, maxy = double.MinValue;
                foreach (Face face in solid.Faces)
                {
                    PlanarFace planarFace = face as PlanarFace;
                    if (planarFace.Origin.Z < min)
                    {
                        min = planarFace.Origin.Z;
                    }
                    else if (planarFace.Origin.Z > max)
                    {
                        max = planarFace.Origin.Z;
                    }
                    if (planarFace.Origin.X < minx)
                    {
                        minx = planarFace.Origin.X;
                    }
                    else if (planarFace.Origin.X > maxx)
                    {
                        maxx = planarFace.Origin.X;
                    }
                    if (planarFace.Origin.Y < miny)
                    {
                        miny = planarFace.Origin.Y;
                    }
                    else if (planarFace.Origin.Y > maxy)
                    {
                        maxy = planarFace.Origin.Y;
                    }
                }

                PlanarFace bottomFace = solid.Faces.get_Item(1) as PlanarFace;
                EdgeArrayArray edgeArrArr = bottomFace.EdgeLoops;
                double x = 0, y = 0;
                foreach (EdgeArray edgeArray in edgeArrArr)
                {
                    for (int i = 0; i < edgeArray.Size; i++)
                    {
                        Line c1 = edgeArray.get_Item(i).AsCurve() as Line;
                        if (c1.Length < 150 / 304.8) continue;
                        Line curve = Line.CreateBound(c1.GetEndPoint(1), c1.GetEndPoint(0));
                        BoundingBoxXYZ bx = new BoundingBoxXYZ();
                        XYZ viewDir = curve.Direction.CrossProduct(XYZ.BasisZ).Normalize();
                        XYZ startPt = curve.GetEndPoint(0);
                        XYZ endPt = curve.GetEndPoint(1).Add(20 * viewDir);
                        XYZ CG = curve.Evaluate(0.5, false);
                        XYZ crossDir = curve.Direction.Normalize().CrossProduct(XYZ.BasisZ).Normalize();
                        XYZ p1 = CG.Add(2 * crossDir);
                        XYZ p2 = CG.Add(-2 * crossDir);
                        double h = max - min;
                        XYZ bxMin = new XYZ(-1.5, 0, -0.50);
                        XYZ bxMax = new XYZ(curve.Length + 1.5, h, 1.5);
                        Transform transform = Transform.Identity;
                        transform.Origin = CG;
                        transform.BasisX = curve.Direction;
                        transform.BasisY = XYZ.BasisZ;
                        bx.Min = bxMin;
                        bx.Max = bxMax;
                        transform.BasisZ = viewDir;
                        bx.Transform = transform;
                        try
                        {
                            ViewSection section = ViewSection.CreateSection(doc, viewForSection.Id, bx);
                            if (UI.sectionVTId != null)
                            {
                                section.ViewTemplateId = UI.sectionVTId;
                            }
                            else
                            {
                                section.DetailLevel = ViewDetailLevel.Fine;
                                section.Scale = 25;
                                section.get_Parameter(BuiltInParameter.MODEL_GRAPHICS_STYLE).Set(4);
                            }
                            if (UI.sectionName != null) section.Name = room.Name + " " + UI.sectionName + " " + (i + 1).ToString(); ;
                            viewPorts.Add(section);
                        }
                        catch (Exception e)
                        {
                            sb.AppendLine(e.Message);
                            sb.AppendLine(e.StackTrace);
                        }
                        //break;

                    }
                }
                try
                {

                    BoundingBoxXYZ bxUp = new BoundingBoxXYZ();
                    XYZ minUp = new XYZ(minx - 1.5, miny - 1.5, min + 3);
                    XYZ maxUp = new XYZ(maxx + 1.5, maxy + 1.5, max + 3);
                    bxUp.Min = minUp; bxUp.Max = maxUp;
                    bxUp.Transform = Transform.Identity;
                    ViewSection up = ViewSection.CreateDetail(doc, viewForDetials.Id, bxUp);
                    if (UI.celingVTId != null)
                    {
                        up.ViewTemplateId = UI.celingVTId;
                    }
                    else
                    {
                        up.DetailLevel = ViewDetailLevel.Fine;
                        up.Scale = 25;
                        up.get_Parameter(BuiltInParameter.MODEL_GRAPHICS_STYLE).Set(4);
                    }
                    if (UI.ceilingName != null)
                    {
                        up.Name = room.Name + " " + UI.ceilingName;
                    }
                    viewPorts.Add(up);
                    //ElementTransformUtils.MirrorElement(doc,up.Id,Plane.CreateByNormalAndOrigin(bottomFace.FaceNormal, bottomFace.Origin.Add(5*XYZ.BasisZ)));
                    List<ElementId> tempIds = new List<ElementId>() { UI.floor1VTId, UI.floor2VTId, UI.floor3VTId };
                    List<string> names = new List<string>() { UI.floor1Name, UI.floor2Name, UI.floor3Name };
                    for (int i = 0; i < 3; i++)
                    {
                        BoundingBoxXYZ bxDown = new BoundingBoxXYZ();
                        Transform tDown = Transform.Identity;
                        tDown.Origin = bottomFace.Origin;
                        tDown.BasisX = new XYZ(-1, 0, 0);
                        tDown.BasisY = new XYZ(0, 1, 0);
                        tDown.BasisZ = new XYZ(0, 0, -1);
                        bxDown.Min = new XYZ(-1.5, -1.5, -5);
                        bxDown.Max = new XYZ(maxx - minx + 1.5, maxy - miny + 1.5, +5);
                        bxDown.Transform = tDown;
                        ViewSection down = ViewSection.CreateDetail(doc, viewForDetials.Id, bxDown);
                        if (tempIds[i] != null)
                        {
                            down.ViewTemplateId = tempIds[i];
                        }
                        else
                        {
                            down.DetailLevel = ViewDetailLevel.Fine;
                            down.Scale = 25;
                            down.get_Parameter(BuiltInParameter.MODEL_GRAPHICS_STYLE).Set(4);
                        }
                        if (names[i] != null)
                        {
                            down.Name = room.Name + " " + names[i];
                        }
                        viewPorts.Add(down);
                    }
                    double rowHeight = 0;
                    for (int i = 0; i < viewPorts.Count; i++)
                    {
                        View section = viewPorts[i];
                        if ((x + (section.CropBox.Max.X - section.CropBox.Min.X) / 20) < 1000 / 304.8)
                        {
                            if (rowHeight > 0)
                            {
                                y = rowHeight;
                            }
                            else
                            {
                                y = (section.CropBox.Max.Y - section.CropBox.Min.Y) / 40;
                                y += 50 / 304.8;
                            }
                            x += (section.CropBox.Max.X - section.CropBox.Min.X) / 40;
                        }
                        else
                        {
                            x = 50 / 304.8;
                            x += (section.CropBox.Max.X - section.CropBox.Min.X) / 40;
                            y += (section.CropBox.Max.Y - section.CropBox.Min.Y) / 40;
                            y += 50 / 304.8;
                            rowHeight = y;

                        }
                        Viewport vp = Viewport.Create(doc, sheet.Id, section.Id, new XYZ(x, y, 0));
                        x += (section.CropBox.Max.X - section.CropBox.Min.X) / 40;
                        y += (section.CropBox.Max.Y - section.CropBox.Min.Y) / 40;
                    }
                }
                catch (Exception ex)
                {
                    sb.AppendLine(ex.Message);
                    sb.AppendLine(ex.StackTrace);
                }


                if (sb.ToString().Trim().Length > 0) td(sb.ToString());
            }
            tr.Commit();
            return Result.Succeeded;

        }

        void td(string Message)
        {
            TaskDialog.Show("Message", Message);
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
