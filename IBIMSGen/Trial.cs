using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace IBIMSGen
{
    [TransactionAttribute(TransactionMode.Manual)]
    internal class Trial : IExternalCommand
    {
        UIDocument uidoc;
        Document doc;
        StringBuilder sb;
        Options options;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            TaskDialog.Show("Err", "here");
            options = new Options();
            options.ComputeReferences = true;
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;
            sb = new StringBuilder();
            List<RevitLinkInstance> links = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();
            List<Solid> solids = new List<Solid>();
            Transaction tr = new Transaction(doc);
            tr.Start("Draw");
            foreach (RevitLinkInstance rli in links)
            {
                Document linkdoc = rli.GetLinkDocument();
                var floors = new FilteredElementCollector(linkdoc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToList();
                foreach (Element floor in floors)
                {
                    Solid s = getSolid(floor);
                    if (s != null)
                    {

                        try
                        {
                            DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel)).SetShape(new List<GeometryObject> { s });
                        }
                        catch { }
                    }
                }
                //sb.AppendLine(floors.Count.ToString());
            }
            TaskDialog.Show("Err", sb.ToString());
            tr.Commit();
            tr.Dispose();
            //TaskDialog.Show("Info", sb.ToString());
            return Result.Succeeded;
        }

        void getGrid()
        {
            Element selectedElement = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element));
            sb.AppendLine(selectedElement.Name);
            var material = doc.GetElement(selectedElement.GetMaterialIds(false).ElementAt(1)) as Material;
            sb.AppendLine(material.Name);

            FilteredElementCollector fec = new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement));
            var patternElementId = material.SurfaceForegroundPatternId;
            var patternElement = fec.Where(x => x.Id == patternElementId).FirstOrDefault();
            var pattern = patternElement as FillPatternElement;
            sb.AppendLine(pattern?.Name);
            var fillPattern = pattern.GetFillPattern().GetFillGrid(1);
            var patOrigin = fillPattern.Origin;
            sb.AppendLine(patOrigin.ToString());
            var offset = fillPattern.Offset;
            sb.AppendLine(offset.ToString());
            var segments = fillPattern.GetSegments();
            sb.AppendLine(segments.Count.ToString());
            try
            {
            }
            catch
            {

            }
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
