using Office = Microsoft.Office.Core;

namespace Pricop.PowerPointTools
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e) { }
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
