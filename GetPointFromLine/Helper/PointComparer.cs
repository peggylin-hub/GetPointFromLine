using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TunnalCal.Helper
{
    class PointComparer : IEqualityComparer<Point3d>
    {
        public int RoundingDigits { get; set; }
        public PointComparer(int roundingDigits)
        {
            RoundingDigits = roundingDigits;
        }

        public bool Equals(Point3d x, Point3d y)
        {
            return Math.Round(x.X, RoundingDigits, MidpointRounding.AwayFromZero) == Math.Round(y.X, RoundingDigits, MidpointRounding.AwayFromZero) &&
                   Math.Round(x.Y, RoundingDigits, MidpointRounding.AwayFromZero) == Math.Round(y.Y, RoundingDigits, MidpointRounding.AwayFromZero) &&
                   Math.Round(x.Z, RoundingDigits, MidpointRounding.AwayFromZero) == Math.Round(y.Z, RoundingDigits, MidpointRounding.AwayFromZero);
        }

        public int GetHashCode(Point3d obj)
        {
            return Math.Round(obj.X, RoundingDigits, MidpointRounding.AwayFromZero).GetHashCode() ^
                   Math.Round(obj.Y, RoundingDigits, MidpointRounding.AwayFromZero).GetHashCode() ^
                   Math.Round(obj.Z, RoundingDigits, MidpointRounding.AwayFromZero).GetHashCode();
        }
    }
}