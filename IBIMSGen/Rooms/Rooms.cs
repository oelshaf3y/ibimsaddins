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
using Forms = System.Windows.Forms;


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
        Element viewForSection, viewForCeiling, viewForFlooring;
        ElementId titleBlockId;
        FilteredElementCollector views, titleBlocks, vft, levels;
        List<ElementId> viewTempsIds;
        List<View> ceilings, floors;
        View VFF, VFC;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            options = new Options();
            options.ComputeReferences = true;
            viewForSection = null;
            views = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views);
            vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType));
            viewTempsIds = views.Cast<View>().Where(x => x.IsTemplate).OrderBy(x => x.Name).Select(x => x.Id).Distinct().ToList();


            titleBlocks = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).OfClass(typeof(FamilyInstance));
            levels = new FilteredElementCollector(doc).OfClass(typeof(Level));
            ceilings = new List<View>();
            ceilings = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views)
                .Cast<View>().Where(l => l.ViewType == ViewType.CeilingPlan)?.ToList();

            floors = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views)
                .Cast<View>().Where(l => l.ViewType == ViewType.FloorPlan)?.ToList();

            if (ceilings == null) { return Result.Failed; }
            viewForSection = vft.Where(x => ((ViewFamilyType)x).ViewFamily == ViewFamily.Section).FirstOrDefault();
            viewForCeiling = vft.Where(l => (l as ViewFamilyType).ViewFamily == ViewFamily.CeilingPlan).First();
            viewForFlooring = vft.Where(l => (l as ViewFamilyType).ViewFamily == ViewFamily.FloorPlan).FirstOrDefault();

            UI = new RoomsUI(doc, viewTempsIds, titleBlocks.ToList());
            UI.ShowDialog();
            titleBlockId = UI.titleBlockId;
            if (UI.DialogResult == System.Windows.Forms.DialogResult.Cancel) return Result.Failed;

            try
            {
                rooms = uidoc.Selection.PickObjects(ObjectType.Element, new RoomSelectionFilter(), "select rooms").Select(x => doc.GetElement(x) as Room).ToList();
            }
            catch
            {
                return Result.Failed;
            }

            sb = new StringBuilder();
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
                if (solid == null || solid.Faces.IsEmpty)
                {
                    sb.AppendLine($"{room.Name} boundry is not correct.");
                    continue;
                }

                double minX = solid.Faces.Cast<PlanarFace>().OrderBy(face => face.Origin.X).Select(face => face.Origin.X).First();
                double minY = solid.Faces.Cast<PlanarFace>().OrderBy(face => face.Origin.Y).Select(face => face.Origin.Y).First();
                double maxX = solid.Faces.Cast<PlanarFace>().OrderByDescending(face => face.Origin.X).Select(face => face.Origin.X).First();
                double maxY = solid.Faces.Cast<PlanarFace>().OrderByDescending(face => face.Origin.Y).Select(face => face.Origin.Y).First();
                double minZ = solid.Faces.Cast<PlanarFace>().OrderBy(face => face.Origin.Z).Select(face => face.Origin.Z).First();
                double maxZ = solid.Faces.Cast<PlanarFace>().OrderByDescending(face => face.Origin.Z).Select(face => face.Origin.Z).First();

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
                        double h = maxZ - minZ;
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
                    XYZ p1 = new XYZ(minX, minY, minZ);
                    XYZ p2 = new XYZ(maxX, maxY, maxZ);
                    Line line = Line.CreateBound(p1, p2);
                    XYZ mid = line.Evaluate(0.5, true);
                    Level topLevel = levels.Cast<Level>().Where(lev => lev.ProjectElevation <= maxZ).OrderByDescending(l => l.ProjectElevation).FirstOrDefault();
                    Level botLevel = levels.Cast<Level>().Where(lev => lev.ProjectElevation >= minZ).OrderBy(l => l.ProjectElevation).FirstOrDefault();
                    List<View> vTop = null, vBot = null; ;
                    if (topLevel != null)
                    {
                        vTop = ceilings.Where(v => v.GenLevel?.Elevation == topLevel.Elevation).ToList();
                        if (vTop.Count > 1)
                        {
                            ParentView pv = new ParentView(vTop, room.Name, "Ceiling");
                            pv.ShowDialog();
                            if (pv.DialogResult == Forms.DialogResult.Cancel) { return Result.Cancelled; }
                            else { VFC = pv.views.ElementAt(pv.comboBox1.SelectedIndex); }
                        }
                        else
                        {
                            VFC = ViewPlan.Create(doc, viewForCeiling.Id, topLevel.Id);
                        }
                        BoundingBoxXYZ bx = new BoundingBoxXYZ();
                        bx.Min = new XYZ(minX - 1.5, minY - 1.5, minZ + 3.5);
                        bx.Max = new XYZ(maxX + 1.50, maxY + 1.5, maxZ + 5);
                        bx.Transform = Transform.Identity;
                        VFC.CropBox = bx;
                        VFC.CropBoxActive = true;
                        VFC.CropBoxVisible = true;
                        View up = VFC;
                        //View up = ViewSection.CreateCallout(doc, VFC.Id, viewForCeiling.Id, new XYZ(minX - 1.5, minY - 1.5, minZ + 3.5), new XYZ(maxX + 1.50, maxY + 1.5, maxZ + 5));
                        BoundingBoxXYZ bxUp = new BoundingBoxXYZ();
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

                    }
                    //ElementTransformUtils.MirrorElement(doc,up.Id,Plane.CreateByNormalAndOrigin(bottomFace.FaceNormal, bottomFace.Origin.Add(5*XYZ.BasisZ)));
                    List<ElementId> tempIds = new List<ElementId>() { UI.floor1VTId, UI.floor2VTId, UI.floor3VTId };
                    List<string> names = new List<string>() { UI.floor1Name, UI.floor2Name, UI.floor3Name };
                    if (botLevel != null)
                    {

                        vBot = floors.Where(v => v.GenLevel?.Elevation == botLevel.Elevation).ToList();
                        if (vBot.Count > 1)
                        {
                            ParentView pv = new ParentView(vBot, room.Name, "Floor");
                            pv.ShowDialog();
                            if (pv.DialogResult == Forms.DialogResult.Cancel) { return Result.Cancelled; }
                            else { VFF = pv.views.ElementAt(pv.comboBox1.SelectedIndex); }
                        }
                        else
                        {
                            VFF = ViewPlan.Create(doc, viewForFlooring.Id, botLevel.Id);
                        }
                        BoundingBoxXYZ bxd = new BoundingBoxXYZ();
                        bxd.Min = new XYZ(minX - 1.5, minY - 1.5, 0);
                        bxd.Max = new XYZ(maxX + 1.5, maxY + 1.5, 0);
                        bxd.Transform = Transform.Identity;
                        VFF.CropBox = bxd;
                        VFF.CropBoxActive = true;
                        VFF.CropBoxVisible = true;
                        for (int i = 0; i < 3; i++)
                        {
                            ElementId downId = VFF.Duplicate(ViewDuplicateOption.AsDependent);
                            View down = doc.GetElement(downId) as View;
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



            }
            tr.Commit();
            if (sb.ToString().Trim().Length > 0) td(sb.ToString());
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
