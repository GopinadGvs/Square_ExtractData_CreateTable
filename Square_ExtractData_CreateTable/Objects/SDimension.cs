using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace Square_ExtractData_CreateTable
{
    public class SDimension
    {
        public Point3d position;
        public string Text;
        public Dimension dimension;

        public SDimension(Dimension dimension,string Text, Point3d position)
        {
            this.dimension = dimension;
            this.position = position;
            this.Text = Text;
        }
    }
}
