using Office = Microsoft.Office.Core;

namespace Pricop.PowerPointTools
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Stream Deck no longer depends on an in-process bridge.
            // The Ribbon remains a normal VSTO add-in; Stream Deck uses its own
            // direct COM helper for maximum reliability.
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }

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
