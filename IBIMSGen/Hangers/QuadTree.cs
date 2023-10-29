using Autodesk.Revit.DB;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IBIMSGen.Hangers
{
    public class QuadTree
    {
        int Capacity = 4;
        Boundary Boundary;
        List<Element> elems;
        QuadTree NorthEastUp, NorthWestUp, SouthEastUp, SouthWestUp, NorthEastDown, NorthWestDown, SouthEastDown, SouthWestDown;
        bool subdevided = false;
        double Top, Left, Right, Bottom, Up, Down;
        public QuadTree(double left, double top, double right, double bottom, double up, double down)
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
            Up = up;
            Down = down;
            elems = new List<Element>();
            Boundary = new Boundary(Left, Top, Right, Bottom, Up, Down);

        }

        public bool insert(Element elem)
        {
            if (!Boundary.contains(elem)) return false;
            if (elems.Count < Capacity && !subdevided)
            {
                elems.Add(elem);
                return true;
            }
            if (!subdevided) subdevide();
            if (NorthEastUp.insert(elem)) return true;
            if (NorthEastDown.insert(elem)) return true;
            if (NorthWestUp.insert(elem)) return true;
            if (NorthWestDown.insert(elem)) return true;
            if (SouthEastUp.insert(elem)) return true;
            if (SouthEastDown.insert(elem)) return true;
            if (SouthWestUp.insert(elem)) return true;
            if (SouthWestDown.insert(elem)) return true;
            return false;
        }

        private void subdevide()
        {
            subdevided = true;
            double midx = (Right + Left) / 2;
            double midy = (Top + Bottom) / 2;
            double midz = (Up + Down) / 2;
            NorthEastUp = new QuadTree(midx, Top, Right, midy, Up, midz);
            NorthEastDown = new QuadTree(midx, Top, Right, midy, midz, Down);
            NorthWestUp = new QuadTree(Left, Top, midx, midy, Up, midz);
            NorthWestDown = new QuadTree(Left, Top, midx, midy, midz, Down);
            SouthEastUp = new QuadTree(midx, midy, Right, Down, Up, midz);
            SouthEastDown = new QuadTree(midx, midy, Right, Down, midz, Down);
            SouthWestUp = new QuadTree(Left, midy, midx, Bottom, Up, midz);
            SouthWestDown = new QuadTree(Left, midy, midx, Bottom, midz, Down);
        }



        public List<Element> query(Boundary range)
        {
            List<Element> result = new List<Element>();
            if (!Boundary.intersects(range)) return result;
            foreach (Element elem in elems)
            {
                if (range.contains(elem)) { result.Add(elem); }
            }
            if (subdevided)
            {
                result.AddRange(NorthEastUp.query(range).ToArray());
                result.AddRange(NorthEastDown.query(range).ToArray());
                result.AddRange(NorthWestUp.query(range).ToArray());
                result.AddRange(NorthWestDown.query(range).ToArray());
                result.AddRange(SouthEastUp.query(range).ToArray());
                result.AddRange(SouthEastDown.query(range).ToArray());
                result.AddRange(SouthWestUp.query(range).ToArray());
                result.AddRange(SouthWestDown.query(range).ToArray());
            }
            return result;
        }
    }

    public class Boundary
    {
        public double Top, Left, Right, Bottom, Up, Down;
        public Boundary(double left, double top, double right, double bottom, double up, double down)
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
            Up = up;
            Down = down;
        }
        public bool intersects(Boundary other)
        {
            return !(Top < other.Bottom || Bottom > other.Top || Left > other.Right || Right < other.Left || Up < other.Down || Down > other.Up);
        }
        public bool contains(Element elem)
        {
            List<XYZ> points = getLocation(elem);
            return points.Where(point => (point.X < Right && point.X > Left && point.Y < Top && point.Y > Bottom && point.Z < Up && point.Z > Down)).Any();
        }
        private List<XYZ> getLocation(Element elem)
        {
            Location location = elem.Location;
            if (location is LocationPoint)
            {
                return new List<XYZ> { ((LocationPoint)location).Point };
            }
            else
            {
                Curve curve = ((LocationCurve)location).Curve;
                return new List<XYZ> { curve.Evaluate(0, true), curve.Evaluate(0.5, true), curve.Evaluate(1, true) };
            }
        }
    }

}
