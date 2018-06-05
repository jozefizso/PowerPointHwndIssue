using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Tools;
using NetOffice.Tools;

namespace PowerPointNetOfficeAddin
{
    [COMAddin("PowerPoint HWND Addin", "Sample add-in to test HWND functionality", 3)]
    [ProgId("PowerPointNetOfficeAddin.Connect")]
    [Guid("054725D8-1D34-4AA9-A8F4-91B3D53D4277")]
    [ComVisible(true)]
    public class PowerPointAddin : COMAddin
    {
        public PowerPointAddin()
        {
            this.OnConnection += PowerPointAddin_OnConnection;
        }

        private void PowerPointAddin_OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            this.Application.NewPresentationEvent += Application_NewPresentation;
        }

        private void Application_NewPresentation(Presentation presentation)
        {
            try
            {
                foreach (DocumentWindow documentWindow in presentation.Windows)
                {
                    var windowHandle = documentWindow.HWND;
                    var caption = documentWindow.Caption;
                    MessageBox.Show($"PowerPoint window handle = {windowHandle}. Caption = {caption}", "PowerPoint NetOffice Add-in", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show($"Failed to retrieve window handle. Exception: {exception.Message}", "PowerPoint NetOffice Add-in", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
