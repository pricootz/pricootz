using Office = Microsoft.Office.Core;

namespace Pricop.PowerPointTools
{
    public partial class ThisAddIn
    {
        private PipeCommandServer _pipeServer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _pipeServer = new PipeCommandServer();
            _pipeServer.Start();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (_pipeServer != null)
            {
                _pipeServer.Dispose();
                _pipeServer = null;
            }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new PricopRibbon();
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
