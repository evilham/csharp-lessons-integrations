using System;
using System.Collections.Generic;
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
            // Write changed address in Cell B2.
            // Careful: Indexes start with 1 in Excel!
            Sheet.Cells[2, 2].Value = string.Format(
                "Cell {0} was changed", Target.Address);
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
