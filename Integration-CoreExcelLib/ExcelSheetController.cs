using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

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
        }

        /// <summary>
        /// This method is called each time the sheet changes.
        /// Use with caution.
        /// </summary>
        /// <param name="Target">Changed range</param>
        private void Sheet_Change(Excel.Range Target)
        {
            // Let's add Cells on Column B, starting on B6 and stopping on
            // first empty cell
            if (Target.Row >= 6)
            {
                Debug.WriteLine("A Cell was changed on a row after the sixth!");
                int curRow = 6; // Start on Row 6
                double sum = 0.0;
                // As long as the value on this row and column B is not empty
                while (!string.IsNullOrWhiteSpace(Sheet.Cells[curRow, 2].Text))
                {
                    // Check the value
                    double value = 0.0;
                    try
                    {
                        value = (double)Sheet.Cells[curRow, 2].Value;
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    { }
                    // Add value (or 0.0 if it wasn't a number)
                    sum += value;
                    // Keep iterating with next row
                    ++curRow;
                }
                // Put value on Cel B2
                Sheet.Cells[2, 2].Value = sum;
            }
        }

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
