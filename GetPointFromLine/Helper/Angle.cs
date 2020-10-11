using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TunnalCal.Helper
{
    class Angle
    {
        public static double RadtoAng(double rad)
        {
            double ang = rad * (180 / Math.PI);
            return ang;
        }

        public static double angToRad(double ang)
        {
            double rad = ang * Math.PI / 180;
            return rad;
        }

    }
}
