using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PolygonData
{

    /// <summary>
    /// Simple 2D Polygon class.
    /// This class must be public so that we can use it from other projects.
    /// </summary>
    public class AcadPolygon
    {
        /// <summary>
        /// A basic 2D Point
        /// </summary>
        public class Point
        {
            public double X = 0.0;
            public double Y = 0.0;
        }
        /// <summary>
        /// This name will be used when we can't find a better one.
        /// </summary>
        private static string defaultName = "Polygon";
        /// <summary>
        /// Dictionary converting number of vertices to polygon name.
        /// </summary>
        private static Dictionary<int, string> polygonNames = new Dictionary<int, string>()
        {
            {3, "Triangle" },
            {4, "Rectangle" },
            {5, "Pentagon" },
            {6, "Hexagon" },
            {7, "Heptagon" },
            {8, "Octagon" },
            {9, "Nonagon" },
            {10, "Decagon" },
            {11, "Hendecagon" },
            {12, "Dodecagon" },
        };
        /// <summary>
        /// Return the name of a polygon with a given amount of vertices.
        /// </summary>
        /// <param name="vertices">Amount of vertices</param>
        /// <returns>Name of the polygon</returns>
        private static string PolygonName(int vertices)
        {
            // Check the dictionary and return that value
            if (polygonNames.ContainsKey(vertices))
                return polygonNames[vertices];
            // Otherwise return the default value
            return defaultName;
        }

        /// <summary>
        /// Container for the vertices
        /// </summary>
        public List<Point> Points = new List<Point>();
        public string Name
        {
            get
            {
                return PolygonName(Points.Count);
            }
        }
    }
}
