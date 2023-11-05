using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace IBIMSGen.Hangers
{
    internal class PipeHanger : IHanger
    {
        public Document Document { get; }
        public Solid Solid { get; }
        public Element Element { get; }
        public double Up { get; }
        public double Down { get; }
        public double Width { get; private set; }
        public List<Support> Supports { get; private set; } = new List<Support>();
        public XYZ Perpendicular { get; private set; }
        public FamilySymbol FamilySymbol { get; private set; }
        public List<List<Dictionary<string, double>>> Dimensions { get; }
        public List<FamilySymbol> Symbols { get; }
        public double Negligible { get; }
        public double Offset { get; }
        public List<double> Diameters { get; }
        public RevitLinkInstance DocumentRLI { get; }
        public RevitLinkInstance LinkInstance { get; }

        public bool isValid { get; private set; } = false;

        public XYZ startPt { get; private set; }
        public bool isFireFighting { get; private set; }
        public ElementId levelId { get; private set; }
        public double OutsideDiam { get; private set; }
        public double pipeElevation { get; private set; }
        public double midElevStart { get; private set; }
        public double midElevEnd { get; private set; }
        public double slope { get; private set; }
        public XYZ pipeDirection { get; private set; }
        public Face LowerFace { get; private set; }
        public FamilySymbol FamilySymbol2 { get; private set; }
        public List<List<string>> AllLinksNames { get; private set; }
        public double Spacing { get; private set; } = 0;
        public double HangerDiameter { get; private set; }
        public double PipeDiam { get; private set; }

        public PipeHanger(Document document, Solid solid, Element element, double up, double down,
            List<List<Dictionary<string, double>>> dimensions, List<FamilySymbol> symbols, double negligible,
            double offset, List<double> diameters, List<List<string>> linksNames, RevitLinkInstance linkInstance, RevitLinkInstance pipesRLI = null)
        {
            Document = document;
            Solid = solid;
            Element = element;
            Up = up;
            Down = down;
            Dimensions = dimensions;
            Symbols = symbols;
            Negligible = negligible;
            Offset = offset;
            Diameters = diameters;
            AllLinksNames = linksNames;
            DocumentRLI = linkInstance;
            LinkInstance = pipesRLI;
            Process();
        }

        public void Process()
        {
            LowerFace = Solid.Faces.get_Item(0);
            FamilySymbol = null;
            FamilySymbol2 = null;
            Supports = new List<Support>();
            isFireFighting = false;
            try
            {
                Width = Element.LookupParameter("Diameter").AsDouble() * 304.8; //mm
            }
            catch
            {
                return;
            }
            Curve pipeCurve = ((LocationCurve)Element.Location).Curve;
            pipeDirection = ((Line)pipeCurve).Direction.Normalize();
            Perpendicular = new XYZ(-pipeDirection.Y, pipeDirection.X, pipeDirection.Z);
            startPt = pipeCurve.Evaluate(0, true);
            XYZ Pf = pipeCurve.Evaluate(1, true);
            XYZ Ps = startPt.Add(Offset * pipeDirection);
            XYZ Pe = Pf.Add(-Offset * pipeDirection);
            Curve hangCurve = null;
            int Rank;
            try
            {
                hangCurve = Line.CreateBound(Ps, Pe);
            }
            catch
            {
                return;
            }
            if (LinkInstance != null)
            {
                RevitLinkType rlt = Document.GetElement(LinkInstance.GetTypeId()) as RevitLinkType;
                Rank = GetSystemRank(rlt.Name);
                if (Rank == -1) return;
                Spacing = GetSysSpacing(Dimensions[Rank], Width);

                if (Spacing == 0) return;
                string familySymbolName = Symbols.Select(x => x.FamilyName).Distinct().ElementAt(Convert.ToInt32(Dimensions[Rank][0]["family"]));
                FamilySymbol = Symbols.Where(x => x.FamilyName.Equals(familySymbolName)).First();
                isFireFighting = Dimensions[Rank][0]["FF"] == 1;
                if (!isFireFighting)
                {
                    string familySymbolName2 = Symbols.Select(x => x.FamilyName).Distinct().ElementAt(Convert.ToInt32(Dimensions[Rank][0]["family2"]));
                    FamilySymbol2 = Symbols.Where(x => x.FamilyName.Equals(familySymbolName2)).First();
                }

            }
            else
            {
                int r = -1;
                try
                {
                    r = GetSystemRank("used", true).Where(x => x != -1 && x != 0 && x != 5).FirstOrDefault();
                }
                catch
                {
                    return;
                }
                TaskDialog.Show("R", r.ToString());
                if (r == -1 || r==0 || r ==5) return;
                Spacing = GetSysSpacing(Dimensions[r], Width);
                if (Spacing == 0) return;

                isFireFighting = Dimensions[r].First()["FF"] == 1;
                int index = Convert.ToInt32(Dimensions[r].Where(x => x["spacing"] != 0).First()["family"]);
                if (index == -1) return;
                string familySymbolName = Symbols.Select(x => x.FamilyName).Distinct().ElementAt(index);
                FamilySymbol = Symbols.Where(x => x.FamilyName.Equals(familySymbolName)).First();
                int index2 = Convert.ToInt32(Dimensions[r].Where(x => x["spacing"] != 0).First()["family2"]);
                if (index2 >= 0)
                {
                    string familySymbolName2 = Symbols.Select(x => x.FamilyName).Distinct().ElementAt(index2);
                    FamilySymbol2 = Symbols.Where(x => x.FamilyName.Equals(familySymbolName2)).First();
                }
            }

            List<XYZ> pps = new List<XYZ>();
            List<XYZ> pipeHangPts = new List<XYZ>();
            try
            {

                OutsideDiam = Element.LookupParameter("Outside Diameter").AsDouble();
                PipeDiam = (OutsideDiam) + (2 * Element.LookupParameter("Insulation Thickness").AsDouble());
            }
            catch
            {
                PipeDiam = (Width / 304.8) + (2 * Element.LookupParameter("Insulation Thickness").AsDouble());

            }
            HangerDiameter = -1;
            foreach (double hangerDiam in Diameters)
            {
                if (PipeDiam <= (hangerDiam / 304.8))
                {
                    HangerDiameter = hangerDiam / 304.8;
                    break;
                }
            }
            if (HangerDiameter == -1) { HangerDiameter = 402 / 304.8; }
            levelId = Element.LookupParameter("Reference Level").AsElementId();
            pipeElevation = pipeCurve.GetEndPoint(0).Z - Element.LookupParameter("Insulation Thickness").AsDouble();
            midElevStart = Element.LookupParameter("Start Middle Elevation").AsDouble();
            midElevEnd = Element.LookupParameter("End Middle Elevation").AsDouble();
            slope = Element.LookupParameter("Slope").AsDouble();
            if (pipeCurve.Length > Negligible && pipeCurve.Length <= Offset)
            {
                XYZ midPt = pipeCurve.Evaluate(0.50, true);
                if (!pipeHangPts.Where(x => x.IsAlmostEqualTo(midPt)).Any())
                {
                    pipeHangPts.Add(midPt);
                    double rod = GetRod(midPt);
                    if (rod != 0) Supports.Add(new Support(midPt, rod + 2995 / 304.8));
                }
            }
            else if (pipeCurve.Length <= Spacing && pipeCurve.Length > Offset)
            {
                if (!pipeHangPts.Where(x => x.IsAlmostEqualTo(Ps)).Any())
                {
                    pipeHangPts.Add(Ps);
                    double rod = GetRod(Ps);
                    if (rod != 0) Supports.Add(new Support(Ps, rod + 2995 / 304.8));

                }
                if (!pipeHangPts.Where(x => x.IsAlmostEqualTo(Pe)).Any())
                {
                    pipeHangPts.Add(Pe);
                    double rod = GetRod(Pe);
                    if (rod != 0) Supports.Add(new Support(Pe, rod + 2995 / 304.8));

                }
            }
            else if (pipeCurve.Length > Spacing)
            {
                if (!pipeHangPts.Where(x => x.IsAlmostEqualTo(Ps)).Any())
                {
                    pipeHangPts.Add(Ps);
                    double rod = GetRod(Ps);
                    if (rod != 0) Supports.Add(new Support(Ps, rod + 2995 / 304.8));
                }
                double n = Math.Ceiling(hangCurve.Length / Spacing) - 1;
                XYZ prev = Ps;
                double Ns = (hangCurve.Length / (n + 1));
                for (int i = 0; i < n; i++)
                {
                    XYZ point = prev.Add(Ns * pipeDirection);
                    if (!pipeHangPts.Where(x => x.IsAlmostEqualTo(point)).Any())
                    {
                        pipeHangPts.Add(point);
                        double rod = GetRod(point);
                        if (rod != 0) Supports.Add(new Support(point, rod + 2995 / 304.8));
                    }
                    prev = point;
                }
                if (!pipeHangPts.Where(x => x.IsAlmostEqualTo(Pe)).Any())
                {
                    pipeHangPts.Add(Pe);
                    double rod = GetRod(Pe);
                    if (rod != 0) Supports.Add(new Support(Pe, rod + 2995 / 304.8));

                }
            }
            isValid = true;
        }

        public void Plant()
        {
            if (!isValid) return;
            foreach (Support support in Supports)
            {
                if (isFireFighting)
                {
                    FamilySymbol.Activate();
                    FamilyInstance hangerFamInst = Document.Create
                        .NewFamilyInstance(support.point, FamilySymbol, Perpendicular, Document.GetElement(levelId), StructuralType.NonStructural);
                    double q = 1;
                    if (midElevEnd < midElevStart)
                    {
                        q = -1;
                    }
                    double elev = pipeElevation + (q * slope * support.point.DistanceTo(startPt));
                    double pipeOffsetFromHost = elev - (3000 / 304.8);
                    if (hangerFamInst == null) { return; }
                    hangerFamInst.LookupParameter("Diameter").Set(PipeDiam);
                    hangerFamInst.LookupParameter("Offset from Host").Set(pipeOffsetFromHost);
                    hangerFamInst.LookupParameter("AnchorElevation").Set(support.rod);
                }
                else
                {
                    Reference reference = null;
                    if (Math.Round(Math.Abs(((PlanarFace)LowerFace).FaceNormal.Z), 3) == 1)
                    {
                        if (LowerFace.Reference.CreateLinkReference(DocumentRLI) != null)
                        {
                            reference = LowerFace.Reference.CreateLinkReference(DocumentRLI);
                        }
                    }
                    else // Ramp
                    {
                        FilteredElementCollector airs = new FilteredElementCollector(Document).OfClass(typeof(FamilySymbol));
                        FamilySymbol fsair = null;
                        foreach (FamilySymbol fsa in airs)
                        {
                            if (fsa.Name == "Air") { fsair = fsa; fsair.Activate(); }
                        }
                        FamilyInstance air = null;
                        air = Document.Create.NewFamilyInstance(support.point, fsair, StructuralType.NonStructural);
                        Document.Regenerate();
                        Face faceair = getSolid(air).Faces.get_Item(0);
                        string refnew = air.UniqueId + ":0:INSTANCE:" + faceair.Reference.ConvertToStableRepresentation(Document);
                        reference = Reference.ParseFromStableRepresentation(Document, refnew);
                    }
                    try
                    {
                        FamilySymbol fs = null;
                        if (Width > (202))
                        {
                            fs = FamilySymbol2;
                            if (fs != null)
                            {
                                fs.Activate();
                            }
                            else
                            {
                                fs = FamilySymbol;
                                fs.Activate();
                            }
                        }
                        else
                        {
                            fs = FamilySymbol;
                            fs.Activate();
                        }
                        FamilyInstance pang = Document.Create.NewFamilyInstance(reference, support.point, pipeDirection, fs);
                        pang.LookupParameter("Schedule Level").Set(levelId);
                        Line ll = Line.CreateUnbound(support.point, XYZ.BasisZ);
                        double rr = pang.HandOrientation.AngleOnPlaneTo(pipeDirection, XYZ.BasisZ);
                        IntersectionResultArray iraa = new IntersectionResultArray();
                        SetComparisonResult scr = LowerFace.Intersect(ll, out iraa);
                        if (iraa != null && !iraa.IsEmpty)
                        {
                            Curve cv = Line.CreateBound(support.point, iraa.get_Item(0).XYZPoint);
                            pang.LookupParameter("Pipe_distance").Set(cv.Length - (0.5 * PipeDiam));
                        }
                        pang.LookupParameter("Pipe Outer Diameter").Set(PipeDiam);
                    }
                    catch
                    {
                    }
                }
            }
        }

        public double GetRod(XYZ point)
        {
            Line tempLine = Line.CreateUnbound(point, XYZ.BasisZ);
            Face lower = Solid.Faces.get_Item(0);
            //Face upper = Solid.Faces.get_Item(1);
            lower.Intersect(tempLine, out IntersectionResultArray intersectionWithLower);
            //upper.Intersect(tempLine, out IntersectionResultArray intersectionWithUpper);
            if (intersectionWithLower == null || intersectionWithLower.IsEmpty) return 0;
            //if () return 0;
            XYZ ipWithLower = intersectionWithLower.get_Item(0).XYZPoint;
            //XYZ ipWithUpper = intersectionWithUpper.get_Item(0).XYZPoint;
            if (ipWithLower.Z > point.Z)
            {
                return ipWithLower.Z - point.Z;
            }
            else
            {
                return 0;
            }
        }


        public int GetSystemRank(string name)
        {
            for (int i = 0; i < AllLinksNames.Count; i++)
            {
                List<string> linksnames = AllLinksNames[i];
                if (linksnames.Where(x => x == name).Any()) return i;
            }
            return -1;
        }

        public List<int> GetSystemRank(string name, bool getAll)
        {
            List<int> foundRanks = new List<int>();
            for (int i = 0; i < AllLinksNames.Count; i++)
            {
                List<string> linksnames = AllLinksNames[i];
                if (linksnames.Where(x => x == name).Any()) foundRanks.Add(i);
            }
            return foundRanks;
        }

        public double GetSysSpacing(List<Dictionary<string, double>> dimensions, double diameter)
        {
            if (dimensions == null) return 0;
            if (dimensions.Count == 0) return 0;
            else if (dimensions.Count == 1) return dimensions[0]["spacing"] / 304.8;
            else if (dimensions.Where(x => Math.Round(diameter, 5) <= Math.Round(x["size"], 5)).Any())
                return dimensions.Where(x => Math.Round(diameter, 5) <= Math.Round(x["size"], 5)).First()["spacing"] / 304.8;
            else return dimensions.Last()["spacing"] / 304.8;
        }
        public Solid getSolid(Element elem)
        {
            IList<Solid> solids = new List<Solid>();
            Options options = new Options();
            options.ComputeReferences = true;
            try
            {
                GeometryElement geo = elem.get_Geometry(options);
                if (geo.FirstOrDefault() is Solid)
                {
                    return (Solid)geo.FirstOrDefault();
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
                return solids.OrderByDescending(x => x.Volume).ElementAt(0);
            }
            else
            {
                return null;
            }
        }

    }
}