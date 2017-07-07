using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace Integration_CoreExcelLib
{
    /// <summary>
    /// This class must be public, that way we can use it from our
    /// other projects.
    /// </summary>
    public class ExcelWorkbookController : IDisposable
    {
        /// <summary>
        /// This is the Excel Workbook we are working on.
        /// </summary>
        public Excel.Workbook Workbook { get; private set; }

        /// <summary>
        /// This will contain the controller for our active sheet.
        /// </summary>
        private ExcelSheetController ActiveSheet { get; set; }

        /// <summary>
        /// Disable parameterless constructor, we need a Workbook to work on.
        /// </summary>
        private ExcelWorkbookController() { }

        /// <summary>
        /// Create the controller while passing the Workbook to work on.
        /// </summary>
        /// <param name="workbook">Workbook to work on</param>
        public ExcelWorkbookController(Excel.Workbook workbook)
        {
            Workbook = workbook;
            // Add ActiveSheet controller
            ActiveSheet = new ExcelSheetController(Workbook.ActiveSheet);

            // Add event handlers
            Workbook.SheetActivate += Workbook_SheetActivate;
            Workbook.SheetDeactivate += Workbook_SheetDeactivate;
        }

        private void Workbook_SheetDeactivate(object Sh)
        {
            if (ActiveSheet == null) return;
            ActiveSheet.Dispose();
            ActiveSheet = null;
        }

        private void Workbook_SheetActivate(object Sh)
        {
            ActiveSheet = new ExcelSheetController((Excel.Worksheet)Sh);
        }

        /// <summary>
        /// Workbook is closing, perform cleanup.
        /// </summary>
        public void Dispose()
        {
            if (disposed) return;
            disposed = true;

            // Remove event handlers

            Workbook = null;
        }
        private bool disposed = false;
    }
}
