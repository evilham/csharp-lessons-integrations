using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using Integration_CoreExcelLib;

namespace Integration_Excel2013
{
    public partial class ThisWorkbook
    {
        private ExcelWorkbookController WorkbookController;
        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            WorkbookController = new ExcelWorkbookController(InnerObject);
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
            WorkbookController.Dispose();
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
