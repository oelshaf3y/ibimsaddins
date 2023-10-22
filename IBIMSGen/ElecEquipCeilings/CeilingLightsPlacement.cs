using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IBIMSGen.ElecEquipCeilings
{
    [TransactionAttribute(TransactionMode.Manual)]
    internal class CeilingLightsPlacement : IExternalCommand
    {
        UIDocument uidoc;
        Document doc;
        List<FamilySymbol> Symbols;
        FamilySymbol fs;
        Options options;
        PlaceLightsUi ui;
        int nx, ny;
        double pitchX, pitchY;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            Reference refFace;
            try
            {
                refFace = uidoc.Selection.PickObject(ObjectType.PointOnElement, new ceilingFloorSelectionFilter(), "Pick Host");
            }
            catch
            {
                return Result.Cancelled;
            }
            Symbols = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_LightingFixtures).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
            options = new Options();
            options.ComputeReferences = true;
            ui = new PlaceLightsUi(Symbols.ToList());
            ui.ShowDialog();
            if (ui.DialogResult == DialogResult.Cancel) return Result.Cancelled;
            fs = ui.families.ElementAt(ui.comboBox1.SelectedIndex);
            if (fs == null)
            {
                td("Not found FS");
                return Result.Failed;
            }

            if (ui.radioButton1.Checked)
            {
                pitchX = ui.x / 304.8;
                pitchY = ui.y / 304.8;
            }
            else
            {
                nx = ui.nx;
                ny = ui.ny;
            }
            using (Transaction transaction = new Transaction(doc, "Center Element"))
            {

                transaction.Start();

                #region set workplane
                Plane plane = Plane.CreateByNormalAndOrigin(uidoc.ActiveView.ViewDirection, uidoc.ActiveView.Origin);
                SketchPlane sp = SketchPlane.Create(doc, plane);
                uidoc.ActiveView.SketchPlane = sp;
                uidoc.ActiveView.ShowActiveWorkPlane();
                #endregion
                ElementMulticategoryFilter ceilingsAndFloors = new ElementMulticategoryFilter(new BuiltInCategory[] { BuiltInCategory.OST_Floors, BuiltInCategory.OST_Ceilings });
                XYZ p1, p2, p3, p4;
                Line l12, l23;

                try
                {

                    p1 = uidoc.Selection.PickPoint(ObjectSnapTypes.Endpoints, "Pick 1st Corner");
                    p2 = uidoc.Selection.PickPoint(ObjectSnapTypes.Endpoints, "Pick 2nd Corner");
                    p3 = uidoc.Selection.PickPoint(ObjectSnapTypes.Endpoints, "Pick 3rd Corner");
                    l12 = Line.CreateBound(p1, p2);
                    l23 = Line.CreateBound(p2, p3);
                    p4 = p1.Add(l23.Length * l23.Direction);
                }
                catch
                {
                    return Result.Cancelled;
                }
                if (ui.radioButton1.Checked)
                {
                    nx = Convert.ToInt32(Math.Floor(l12.Length / pitchX));
                    ny = Convert.ToInt32(Math.Floor(l23.Length / pitchY));
                }
                else
                {
                    pitchX = l12.Length / nx;
                    pitchY = l23.Length / ny;
                }

                Element elem = doc.GetElement(refFace);
                if (elem is RevitLinkInstance)
                {
                    RevitLinkInstance rli = elem as RevitLinkInstance;
                    Document LinkDoc = rli.GetLinkDocument();
                    elem = LinkDoc.GetElement(refFace.LinkedElementId);
                }
                Solid solid = getSolid(elem);
                double z = solid.Faces.Cast<PlanarFace>().OrderBy(x => x.Origin.Z).Select(x => x.Origin.Z).First();
                fs.Activate();
                for (int i = 0; i < nx; i++)
                {
                    for (int j = 0; j < ny; j++)
                    {
                        XYZ point = p1.Add((pitchX / 2 + pitchX * i) * l12.Direction.Normalize()).Add((pitchY / 2 + pitchY * j) * l23.Direction.Normalize());
                        XYZ cP = new XYZ(point.X, point.Y, z);
                        doc.Create.NewFamilyInstance(refFace, cP, doc.ActiveView.ViewDirection.CrossProduct(XYZ.BasisX).Negate(), fs);
                    }
                }

                transaction.Commit();
                transaction.Dispose();
            }



            return Result.Succeeded;
        }
        void td(string message)
        {
            TaskDialog.Show("Message", message);
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
