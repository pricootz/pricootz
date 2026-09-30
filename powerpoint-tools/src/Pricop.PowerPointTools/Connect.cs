using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Pricop.PowerPointTools
{
    [ComVisible(true)]
    [Guid("1D8EA74E-3C7F-4DA2-BE72-5C4FCF17A061")]
    [ProgId("Pricop.PowerPointTools")]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private static dynamic _app;

        public void OnConnection(object Application, ExtConnectMode ConnectMode, object AddInInst, ref Array custom) => _app = Application;
        public void OnDisconnection(ExtDisconnectMode RemoveMode, ref Array custom) => _app = null;
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public string GetCustomUI(string RibbonID)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var s = asm.GetManifestResourceStream("Pricop.PowerPointTools.Ribbon.xml"))
            using (var r = new StreamReader(s))
                return r.ReadToEnd();
        }

        public void Ribbon_Load(object ribbonUI) { }

        public object GetImage(object control)
        {
            try
            {
                dynamic c = control;
                string id = (string)c.Id;
                switch (id)
                {
                    case "btnAlignLeft": return IconData.GetPicture("left");
                    case "btnAlignCenter": return IconData.GetPicture("center");
                    case "btnAlignRight": return IconData.GetPicture("right");
                    case "btnAlignTop": return IconData.GetPicture("top");
                    case "btnAlignMiddle": return IconData.GetPicture("middle");
                    case "btnAlignBottom": return IconData.GetPicture("bottom");
                    case "btnDistH": return IconData.GetPicture("dist_h");
                    case "btnDistV": return IconData.GetPicture("dist_v");
                    case "btnSameWidth": return IconData.GetPicture("width");
                    case "btnSameHeight": return IconData.GetPicture("height");
                    case "btnSameSize": return IconData.GetPicture("size");
                    case "btnRect": return IconData.GetPicture("rect");
                    case "btnRound": return IconData.GetPicture("round");
                    case "btnShadowOff": return IconData.GetPicture("shadow_off");
                    case "btnShadowOn": return IconData.GetPicture("shadow_on");
                    case "btnLineOff": return IconData.GetPicture("line_off");
                    case "btnLineOn": return IconData.GetPicture("line_on");
                    case "btnMatchStyle": return IconData.GetPicture("copy");
                    case "btnClean": return IconData.GetPicture("clean");
                    default: return null;
                }
            }
            catch { return null; }
        }

        private dynamic Selected(int min = 1)
        {
            if (_app == null) return null;
            dynamic win = _app.ActiveWindow;
            if (win == null) return null;
            dynamic sel = win.Selection;
            if (sel == null || (int)sel.Type != 2) return null;
            dynamic range = sel.ShapeRange;
            return (int)range.Count >= min ? range : null;
        }

        private void Safe(Action action)
        {
            try { action(); }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Pricop Tools", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void AlignLeft(object c) => Safe(() => Selected(2)?.Align(0, 0));
        public void AlignCenter(object c) => Safe(() => Selected(2)?.Align(1, 0));
        public void AlignRight(object c) => Safe(() => Selected(2)?.Align(2, 0));
        public void AlignTop(object c) => Safe(() => Selected(2)?.Align(3, 0));
        public void AlignMiddle(object c) => Safe(() => Selected(2)?.Align(4, 0));
        public void AlignBottom(object c) => Safe(() => Selected(2)?.Align(5, 0));

        public void DistributeH(object c) => Safe(() => Selected(3)?.Distribute(0, 0));
        public void DistributeV(object c) => Safe(() => Selected(3)?.Distribute(1, 0));

        public void SameWidth(object c) => Safe(() =>
        {
            dynamic r = Selected(2); if (r == null) return;
            float v = (float)r[1].Width;
            for (int i = 2; i <= (int)r.Count; i++) r[i].Width = v;
        });

        public void SameHeight(object c) => Safe(() =>
        {
            dynamic r = Selected(2); if (r == null) return;
            float v = (float)r[1].Height;
            for (int i = 2; i <= (int)r.Count; i++) r[i].Height = v;
        });

        public void SameSize(object c) => Safe(() =>
        {
            dynamic r = Selected(2); if (r == null) return;
            float w = (float)r[1].Width, h = (float)r[1].Height;
            for (int i = 2; i <= (int)r.Count; i++) { r[i].Width = w; r[i].Height = h; }
        });

        public void Rectangle(object c) => Safe(() =>
        {
            dynamic r = Selected(); if (r == null) return;
            for (int i = 1; i <= (int)r.Count; i++)
                if ((int)r[i].Type == 1) r[i].AutoShapeType = 1;
        });

        public void RoundedRectangle(object c) => Safe(() =>
        {
            dynamic r = Selected(); if (r == null) return;
            for (int i = 1; i <= (int)r.Count; i++)
                if ((int)r[i].Type == 1) r[i].AutoShapeType = 5;
        });

        public void ShadowOff(object c) => Safe(() =>
        {
            dynamic r = Selected(); if (r == null) return;
            for (int i = 1; i <= (int)r.Count; i++) r[i].Shadow.Visible = 0;
        });

        public void ShadowOn(object c) => Safe(() =>
        {
            dynamic r = Selected(); if (r == null) return;
            for (int i = 1; i <= (int)r.Count; i++)
            {
                dynamic s = r[i].Shadow;
                s.Visible = -1;
                try { s.Type = 21; } catch { }
                try { s.Transparency = 0.55f; } catch { }
                try { s.Blur = 7f; } catch { }
                try { s.OffsetX = 1.5f; s.OffsetY = 2f; } catch { }
            }
        });

        public void LineOff(object c) => Safe(() =>
        {
            dynamic r = Selected(); if (r == null) return;
            for (int i = 1; i <= (int)r.Count; i++) r[i].Line.Visible = 0;
        });

        public void LineOn(object c) => Safe(() =>
        {
            dynamic r = Selected(); if (r == null) return;
            for (int i = 1; i <= (int)r.Count; i++)
            {
                r[i].Line.Visible = -1;
                r[i].Line.Weight = 1f;
            }
        });

        public void MatchStyle(object c) => Safe(() =>
        {
            dynamic r = Selected(2); if (r == null) return;
            r[1].PickUp();
            for (int i = 2; i <= (int)r.Count; i++) r[i].Apply();
        });

        public void CleanBoxes(object c) => Safe(() =>
        {
            dynamic r = Selected(); if (r == null) return;
            float h = (float)r[1].Height;
            for (int i = 1; i <= (int)r.Count; i++)
            {
                if ((int)r[i].Type == 1) r[i].AutoShapeType = 1;
                r[i].Shadow.Visible = 0;
                r[i].Line.Visible = 0;
                if (i > 1) r[i].Height = h;
            }
        });
    }
}
