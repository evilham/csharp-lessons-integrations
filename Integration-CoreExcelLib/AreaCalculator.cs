using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Integration_CoreExcelLib
{
    class AreaCalculator
    {
        public static double CalculateArea(Point2D[] points)
        {
            double[] xValues = points.Select((p) => p.X).ToArray();
            double[] yValues = points.Select((p) => p.Y).ToArray();
            return CalculateArea(xValues, yValues);
        }
        public static double CalculateArea(double[] xValues, double[] yValues)
        {
            double area = 0;
            int count = 0;
            // Go through all vertices and calculate the area.
            while (count < xValues.Length)
            {
                // Add all point pairs.
                if (count < xValues.Length - 1)
                    area += (xValues[count] - xValues[count + 1]) * (yValues[count] + yValues[count + 1]);
                // Don't forget the last pair (last and first point).
                else
                    area += (xValues[count] - xValues[0]) * (yValues[count] + yValues[0]);

                count++;
            }
            // Take the absolute half value.
            return Math.Abs(area / 2);
        }
    }
}
