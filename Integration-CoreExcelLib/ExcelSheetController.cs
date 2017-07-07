using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

using PolygonData;

namespace Integration_CoreExcelLib
{
    /// <summary>
    /// This is where we will program our sample application.
    /// </summary>
    class ExcelSheetController : IDisposable
    {
        /// <summary>
        /// Reference to the sheet, we will use this to interact with it.
        /// </summary>
        private Excel.Worksheet Sheet;

        /// <summary>
        /// Disable parameterless constructor, we need a Worksheet to work on.
        /// </summary>
        private ExcelSheetController() { }

        /// <summary>
        /// Create the controller while passing the Worksheet to work on.
        /// </summary>
        /// <param name="sheet">Sheet to work on</param>
        internal ExcelSheetController(Excel.Worksheet worksheet)
        {
            Sheet = worksheet;

            // Add event handlers
            Sheet.Change += Sheet_Change;

            Startup();
        }

        /// <summary>
        /// To be called when this object is created (Sheet becomes active).
        /// </summary>
        private void Startup()
        {
            // Read polygon data from the server
            string jsonData;
            using (System.Net.WebClient webclient = new System.Net.WebClient())
            {
                jsonData = webclient.DownloadString(
                    "https://evilham.com/polygons.json");
            }

            // Convert json data to .net objects
            List<AcadPolygon> acadPolys = JsonConvert.DeserializeObject<
                List<AcadPolygon>>(jsonData);
            Sheet.Cells[3, 1].Value = string.Format(
                "{0} Polygons",acadPolys.Count);

            // Start writing on line 6
            int curRow = 6;
            // Don't react to changes
            changed_with_code = true;
            // Iterate through polygons read from json
            foreach (AcadPolygon acadPoly in acadPolys)
            {
                // Write name in this row
                Sheet.Cells[curRow, 3].Value = acadPoly.Name;
                int curCol = 4;
                // Write points in this row
                foreach (AcadPolygon.Point acadPoint in acadPoly.Points)
                {
                    Sheet.Cells[curRow, curCol++].Value = acadPoint.X;
                    Sheet.Cells[curRow, curCol++].Value = acadPoint.Y;
                }
                // Go to next row
                ++curRow;
            }
            CodeChanges();
            // React to user changes
            changed_with_code = false;
        }

        /// <summary>
        /// This method is called each time the sheet changes.
        /// Use with caution.
        /// </summary>
        /// <param name="Target">Changed range</param>
        private void Sheet_Change(Excel.Range Target)
        {
            if (changed_with_code) return;
            changed_with_code = true;
            if (Target.Column >= 3 && Target.Row >= 6)
            {
                Debug.WriteLine("A cell was changed in user region");
                CodeChanges();
            }
            changed_with_code = false;
        }
        private void CodeChanges()
        {
            int curRow = 6; // Start on Row 6
                            // As long as the value on this row and column C is not empty
            while (!string.IsNullOrWhiteSpace(Sheet.Cells[curRow, 3].Text))
            {
                string polyType = Sheet.Cells[curRow, 3].Text;
                Debug.WriteLine(polyType);

                // 1. Read points
                List<Point2D> points = new List<Point2D>();
                int curCol = 4; // Start in Column D (1st point)
                                // Keep going as long as there is an X value
                while (!string.IsNullOrWhiteSpace(Sheet.Cells[curRow, curCol].Text))
                {
                    double x = Sheet.Cells[curRow, curCol].Value;
                    // Iterate through columns
                    ++curCol;
                    if (string.IsNullOrWhiteSpace(Sheet.Cells[curRow, curCol].Text))
                        // if Y is not defined, ignore this X value
                        break;
                    double y = Sheet.Cells[curRow, curCol].Value;
                    // Iterate through columns
                    ++curCol;

                    // Create and add point represented by these 2 cells
                    Point2D point = new Point2D(x, y);
                    points.Add(point);
                }
                Debug.WriteLine(string.Format(
                    "{0} points were read for poly {1}.",
                    points.Count, polyType));

                // 2. Create Polygon for this row
                Polygon poly = new Polygon(points.ToArray());

                // 3. Write surface and perimeter
                Sheet.Cells[curRow, 1].Value = poly.Area;

                // Iterate through rows
                ++curRow;
            }
        }
        private bool changed_with_code = false;

        /// <summary>
        /// This controller is not needed any longer, perform cleanup.
        /// </summary>
        public void Dispose()
        {
            if (disposed) return;
            disposed = true;

            // Remove event handlers
            Sheet.Change -= Sheet_Change;

            Sheet = null;
        }
        private bool disposed = false;
    }
}
