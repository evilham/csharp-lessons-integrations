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

            // Write "Hello!" to Cell B2
            Sheet.Cells[2, 2].value = "Hello!";
        }

        /// <summary>
        /// This controller is not needed any longer, perform cleanup.
        /// </summary>
        public void Dispose()
        {
            if (disposed) return;
            disposed = true;

            // Remove event handlers

            Sheet = null;
        }
        private bool disposed = false;
    }
}
