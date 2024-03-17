using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace Highlighter
{

    
    public partial class HighlighterWork
    {

        private static HighlighterWork instance;

        private static string lastSheet;
        private static string lastCell;
        private static double lastFormat;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            instance = this;
            this.Application.SheetSelectionChange += Application_SheetSelectionChange;
        }

        public static void clearLastCell()
        {
            if (lastCell != null)
            {
                Excel.Range last = instance.Application.Worksheets[lastSheet].Range(lastCell);

                if (lastFormat == (double)0XFFFFFF)
                    last.ClearFormats();
                else
                    last.Interior.Color = lastFormat;
            }
        }

        public static void highlightCell()
        {
            Excel.Range range = instance.Application.ActiveCell;

            clearLastCell();

            lastSheet = instance.Application.ActiveSheet.Name;
            lastCell = range.Address;
            lastFormat = range.Interior.Color;

            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(Ribbon.color);
        }

        private void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            if(! Ribbon.toggled)
            {
                return;
            }

            try
            {
                highlightCell();

                Console.WriteLine("Event");
            } catch (Exception ex) {
                Console.WriteLine(ex.ToString());
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
