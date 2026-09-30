#pragma warning disable 414
namespace Pricop.PowerPointTools
{
    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(0)]
    [global::System.Security.Permissions.PermissionSetAttribute(global::System.Security.Permissions.SecurityAction.Demand, Name="FullTrust")]
    public sealed partial class ThisAddIn : Microsoft.Office.Tools.AddInBase
    {
        internal Microsoft.Office.Tools.CustomTaskPaneCollection CustomTaskPanes;
        internal Microsoft.Office.Interop.PowerPoint.Application Application;

        public ThisAddIn(global::Microsoft.Office.Tools.Factory factory, global::System.IServiceProvider serviceProvider)
            : base(factory, serviceProvider, "AddIn", "ThisAddIn")
        {
            Globals.Factory = factory;
        }

        protected override void Initialize()
        {
            base.Initialize();
            this.Application = this.GetHostItem<Microsoft.Office.Interop.PowerPoint.Application>(typeof(Microsoft.Office.Interop.PowerPoint.Application), "Application");
            Globals.ThisAddIn = this;
            global::System.Windows.Forms.Application.EnableVisualStyles();
            this.InitializeControls();
        }

        protected override void FinishInitialization()
        {
            this.InternalStartup();
            this.OnStartup();
        }

        protected override void InitializeDataBindings() { }
        private void InitializeControls()
        {
            this.CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this);
        }

        protected override void OnShutdown()
        {
            this.CustomTaskPanes.Dispose();
            base.OnShutdown();
        }
    }

    internal sealed partial class Globals
    {
        private Globals() { }
        private static ThisAddIn _ThisAddIn;
        private static global::Microsoft.Office.Tools.Factory _factory;

        internal static ThisAddIn ThisAddIn
        {
            get { return _ThisAddIn; }
            set { if (_ThisAddIn == null) _ThisAddIn = value; }
        }

        internal static global::Microsoft.Office.Tools.Factory Factory
        {
            get { return _factory; }
            set { if (_factory == null) _factory = value; }
        }
    }
}
