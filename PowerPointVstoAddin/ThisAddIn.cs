using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointVstoAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ((PowerPoint.EApplication_Event)this.Application).NewPresentation += NewPresentationEventHandler;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void NewPresentationEventHandler(PowerPoint.Presentation presentation)
        {
            try
            {
                foreach (PowerPoint.DocumentWindow documentWindow in presentation.Windows)
                {
                    var windowHandle = documentWindow.HWND;
                    var caption = documentWindow.Caption;
                    MessageBox.Show($"PowerPoint window handle = {windowHandle}. Caption = {caption}", "PowerPoint VSTO Add-in", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show($"Failed to retrieve window handle. Exception: {exception.Message}", "PowerPoint VSTO Add-in", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
