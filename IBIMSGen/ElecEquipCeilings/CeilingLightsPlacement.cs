using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            Symbols = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_LightingFixtures).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
            fs = Symbols.Where(x => x.Name.Equals("CSL-1")).FirstOrDefault();
            options = new Options();
            options.ComputeReferences = true;
            if (fs == null)
            {
                td("Not found FS");
                return Result.Failed;
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
                List<Element> elementInView = new FilteredElementCollector(doc, doc.ActiveView.Id).WherePasses(ceilingsAndFloors).ToList();
                uidoc.Selection.SetElementIds(elementInView.Select(x => x.Id).ToArray());

                XYZ p1 = uidoc.Selection.PickPoint(ObjectSnapTypes.Endpoints, "Pick 1st Corner");
                XYZ p2 = uidoc.Selection.PickPoint(ObjectSnapTypes.Endpoints, "Pick 2nd Corner");
                XYZ p3 = uidoc.Selection.PickPoint(ObjectSnapTypes.Endpoints, "Pick 3rd Corner");
                Line l12 = Line.CreateBound(p1, p2);
                Line l23 = Line.CreateBound(p2, p3);
                XYZ p4 = p1.Add(l23.Length * l23.Direction);
                double nx = 3;
                double ny = 2;
                double pitchX = l12.Length / nx;
                double pitchY = l23.Length / ny;

                Reference refFace = uidoc.Selection.PickObject(ObjectType.Face, "pick Host");
                Element elem = doc.GetElement(refFace);
                var par = elem.get_Parameter(BuiltInParameter.TILE_PATTERN_GRID_CELLS_X);
                td(par.ToString)
                Solid solid = getSolid(elem);
                double z = solid.Faces.Cast<PlanarFace>().OrderBy(x => x.Origin.Z).Select(x => x.Origin.Z).First();

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
