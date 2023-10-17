using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
//using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IBIMSGen.ElecCables
{
    internal class QuadTree
    {

        int Capacity;
        Rectangle Boundary;
        List<XYZ> Points;
        bool Subdevided;
        QuadTree NorthWest, SouthWest, NorthEast, SouthEast;

        public QuadTree( Rectangle boundary)
        {
            Capacity = 4;
            Boundary = boundary;
            Subdevided = false;
            Points = new List<XYZ>();
        }

        public bool Insert(XYZ point)
        {
            if (!this.Contains(point)) return false;
            if (Points.Count < Capacity && !Subdevided)
            {
                if (Points.Where(x => x.IsAlmostEqualTo(point)).Any()) return false;
                Points.Add(point);
                return true;
            }
            else
            {
                if (!Subdevided) Subdevide();
                if (NorthWest.Insert(point)) return true;
                if (NorthEast.Insert(point)) return true;
                if (SouthEast.Insert(point)) return true;
                if (SouthWest.Insert(point)) return true;
                return false;
            }
        }

        private bool Contains(XYZ point)
        {
            if (point.X > Boundary.Right || point.X < Boundary.Left || point.Y < Boundary.Bottom || point.Y > Boundary.Top) return false;
            return true;
        }

        private void Subdevide()
        {
            Subdevided = true;
            int midx = (Boundary.Left + Boundary.Right) / 2;
            int midy = (Boundary.Top + Boundary.Bottom) / 2;
            NorthWest = new QuadTree( new Rectangle(Boundary.Left, Boundary.Top, midx, midy));
            NorthEast = new QuadTree(new Rectangle(midx, Boundary.Top, Boundary.Right, midy));
            SouthEast = new QuadTree(new Rectangle(midx, midy, Boundary.Bottom, Boundary.Right));
            SouthWest = new QuadTree(new Rectangle(Boundary.Left, midy, midx, Boundary.Bottom));
        }
        public List<XYZ> queryRange(QuadTree rect)
        {
            List<XYZ> result = new List<XYZ>();
            if (rect.Boundary.Right < Boundary.Left || rect.Boundary.Left > Boundary.Right || rect.Boundary.Top < Boundary.Bottom || rect.Boundary.Bottom > Boundary.Top) return result;
            for (int i = 0; i < Points.Count; i++)
            {
                if (rect.Contains(Points[i]))
                {
                    result.Add(Points[i]);
                    Points.RemoveAt(i);
                }
            }
            if (!Subdevided) return result;

            result.AddRange(NorthEast.queryRange(rect).ToArray());
            result.AddRange(NorthWest.queryRange(rect).ToArray());
            result.AddRange(SouthEast.queryRange(rect).ToArray());
            result.AddRange(SouthWest.queryRange(rect).ToArray());
            return result;
        }
    }
}
