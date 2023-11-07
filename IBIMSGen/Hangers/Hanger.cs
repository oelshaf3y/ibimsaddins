﻿using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace IBIMSGen.Hangers
{
    public abstract class Hanger
    {
        public Document Document { get; set; }
        public Solid Solid { get; set; }
        public Element Element { get; set; }
        public List<List<Dictionary<string, double>>> Dimensions { get; set; }
        public double Up { get; set; }
        public double Down { get; set; }
        public List<FamilySymbol> Symbols { get; set; }
        public double Negligible { get; set; }
        public double Offset { get; set; }
        public List<Support> Supports { get; set; } = new List<Support>();
        public FamilySymbol FamilySymbol { get; set; }
        public XYZ Perpendicular { get; set; }
        public double Width { get; set; }
        public bool isValid { get; set; } = false;
        public double Spacing { get; set; } = 0;
        public List<Element> ElementsNearby { get; set; } = new List<Element>();
        public RevitLinkInstance LinkInstance { get; set; }
        public QuadTree AllDuctsTree { get; set; }
        public QuadTree AllPipesTree { get; set; }
        public QuadTree AllTraysTree { get; set; }
        public Curve ElementCurve { get; set; }
        public Boundary Region { get; set; }
        public void GetRegion()
        {
            double minx = Math.Min(ElementCurve.GetEndPoint(0).X, ElementCurve.GetEndPoint(1).X) - 5;
            double miny = Math.Min(ElementCurve.GetEndPoint(0).Y, ElementCurve.GetEndPoint(1).Y) - 5;
            double maxx = Math.Max(ElementCurve.GetEndPoint(0).X, ElementCurve.GetEndPoint(1).X) + 5;
            double maxy = Math.Max(ElementCurve.GetEndPoint(0).Y, ElementCurve.GetEndPoint(1).Y) + 5;
            Region = new Boundary(minx, maxy, maxx, miny, Up, Down);
        }
        public double GetRod(XYZ point)
        {

            Line tempLine = Line.CreateUnbound(point, XYZ.BasisZ);
            Face floorLower = Solid.Faces.get_Item(0);
            Face face = null;
            face = GetNearbyElementFaces()?.Where(x => x.Origin.Z > point.Z && x.Origin.Z < ((PlanarFace)floorLower).Origin.Z)?.OrderBy(x => x.Origin.Z)?.FirstOrDefault();
            if (face != null)
            {
                TaskDialog.Show("notNull", "notNull");
                face.Intersect(tempLine, out IntersectionResultArray intersectionWithFace);
                if (intersectionWithFace != null && !intersectionWithFace.IsEmpty)
                {
                    XYZ ipWithFace = intersectionWithFace.get_Item(0).XYZPoint;
                    return ipWithFace.Z - point.Z;
                }
            }
            //Face upper = Solid.Faces.get_Item(1);
            floorLower.Intersect(tempLine, out IntersectionResultArray intersectionWithLower);
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
        public List<XYZ> DecOrder(List<XYZ> points, Curve curve)
        {
            if (Math.Round(((Line)curve).Direction.Normalize().Y, 3) == 0)
            {
                return points.OrderByDescending(a => a.X).ToList();
            }
            else
            {
                return points.OrderByDescending(a => a.Y).ToList();
            }
        }


        public int GetSystemRank(string name, List<List<string>> AllLinksNames)
        {
            for (int i = 0; i < AllLinksNames.Count; i++)
            {
                List<string> linksnames = AllLinksNames[i];
                if (linksnames.Where(x => x == name).Any()) return i;
            }
            return -1;
        }

        public List<int> GetSystemRank(string name, bool getAll, List<List<string>> AllLinksNames)
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
        public Solid GetSolid(Element elem)
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


        public List<PlanarFace> GetNearbyElementFaces()
        {
            ElementsNearby.AddRange(AllDuctsTree.query(Region).ToArray());
            ElementsNearby.AddRange(AllPipesTree.query(Region).ToArray());
            ElementsNearby.AddRange(AllTraysTree.query(Region).ToArray());
            List<PlanarFace> lowerFaces = new List<PlanarFace>();
            foreach (Element elem in ElementsNearby)
            {
                Solid _solid = GetSolid(elem);
                try
                {
                    TaskDialog.Show("Here", "here");
                    Face lowerFace = _solid.Faces.get_Item(0);
                    PlanarFace planar = lowerFace as PlanarFace;
                    lowerFaces.Add(planar);
                }
                catch { }
            }

            return null;
        }
    }
}
