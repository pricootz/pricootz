using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Pricop.PowerPointTools
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public sealed class PricopRibbon : Office.IRibbonExtensibility
    {
        public string GetCustomUI(string ribbonID)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream("Pricop.PowerPointTools.PricopRibbon.xml"))
            using (var reader = new StreamReader(stream))
                return reader.ReadToEnd();
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) { }

        public object GetImage(Office.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
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

        private PowerPoint.ShapeRange Selected(int minimum = 1)
        {
            var app = Globals.ThisAddIn.Application;
            if (app == null || app.ActiveWindow == null) return null;
            var selection = app.ActiveWindow.Selection;
            if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes) return null;
            var range = selection.ShapeRange;
            return range.Count >= minimum ? range : null;
        }

        private void Run(Action action, int minimum)
        {
            try
            {
                var r = Selected(minimum);
                if (r == null)
                {
                    MessageBox.Show(minimum <= 1 ? "Seleziona almeno una forma." : "Seleziona almeno " + minimum + " elementi.",
                        "Pricop Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                action();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Pricop Tools", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void AlignLeft(Office.IRibbonControl c) => Run(() => Selected(2).Align(Office.MsoAlignCmd.msoAlignLefts, Office.MsoTriState.msoFalse), 2);
        public void AlignCenter(Office.IRibbonControl c) => Run(() => Selected(2).Align(Office.MsoAlignCmd.msoAlignCenters, Office.MsoTriState.msoFalse), 2);
        public void AlignRight(Office.IRibbonControl c) => Run(() => Selected(2).Align(Office.MsoAlignCmd.msoAlignRights, Office.MsoTriState.msoFalse), 2);
        public void AlignTop(Office.IRibbonControl c) => Run(() => Selected(2).Align(Office.MsoAlignCmd.msoAlignTops, Office.MsoTriState.msoFalse), 2);
        public void AlignMiddle(Office.IRibbonControl c) => Run(() => Selected(2).Align(Office.MsoAlignCmd.msoAlignMiddles, Office.MsoTriState.msoFalse), 2);
        public void AlignBottom(Office.IRibbonControl c) => Run(() => Selected(2).Align(Office.MsoAlignCmd.msoAlignBottoms, Office.MsoTriState.msoFalse), 2);

        public void DistributeH(Office.IRibbonControl c) => Run(() => Selected(3).Distribute(Office.MsoDistributeCmd.msoDistributeHorizontally, Office.MsoTriState.msoFalse), 3);
        public void DistributeV(Office.IRibbonControl c) => Run(() => Selected(3).Distribute(Office.MsoDistributeCmd.msoDistributeVertically, Office.MsoTriState.msoFalse), 3);

        public void SameWidth(Office.IRibbonControl c) => Run(() =>
        {
            var r = Selected(2);
            float value = r[1].Width;
            for (int i = 2; i <= r.Count; i++) r[i].Width = value;
        }, 2);

        public void SameHeight(Office.IRibbonControl c) => Run(() =>
        {
            var r = Selected(2);
            float value = r[1].Height;
            for (int i = 2; i <= r.Count; i++) r[i].Height = value;
        }, 2);

        public void SameSize(Office.IRibbonControl c) => Run(() =>
        {
            var r = Selected(2);
            float w = r[1].Width, h = r[1].Height;
            for (int i = 2; i <= r.Count; i++) { r[i].Width = w; r[i].Height = h; }
        }, 2);

        public void Rectangle(Office.IRibbonControl c) => Run(() =>
        {
            var r = Selected();
            for (int i = 1; i <= r.Count; i++)
                if (r[i].Type == Office.MsoShapeType.msoAutoShape)
                    r[i].AutoShapeType = Office.MsoAutoShapeType.msoShapeRectangle;
        }, 1);

        public void RoundedRectangle(Office.IRibbonControl c) => Run(() =>
        {
            var r = Selected();
            for (int i = 1; i <= r.Count; i++)
                if (r[i].Type == Office.MsoShapeType.msoAutoShape)
                    r[i].AutoShapeType = Office.MsoAutoShapeType.msoShapeRoundedRectangle;
        }, 1);

        public void ShadowOff(Office.IRibbonControl c) => Run(() =>
        {
            var r = Selected();
            for (int i = 1; i <= r.Count; i++) r[i].Shadow.Visible = Office.MsoTriState.msoFalse;
        }, 1);

        public void ShadowOn(Office.IRibbonControl c) => Run(() =>
        {
            var r = Selected();
            for (int i = 1; i <= r.Count; i++)
            {
                var s = r[i].Shadow;
                s.Visible = Office.MsoTriState.msoTrue;
                s.Transparency = 0.55f;
                s.Blur = 7f;
                s.OffsetX = 1.5f;
                s.OffsetY = 2f;
            }
        }, 1);

        public void LineOff(Office.IRibbonControl c) => Run(() =>
        {
            var r = Selected();
            for (int i = 1; i <= r.Count; i++) r[i].Line.Visible = Office.MsoTriState.msoFalse;
        }, 1);

        public void LineOn(Office.IRibbonControl c) => Run(() =>
        {
            var r = Selected();
            for (int i = 1; i <= r.Count; i++)
            {
                r[i].Line.Visible = Office.MsoTriState.msoTrue;
                r[i].Line.Weight = 1f;
            }
        }, 1);

        public void MatchStyle(Office.IRibbonControl c) => Run(() =>
        {
            var r = Selected(2);
            r[1].PickUp();
            for (int i = 2; i <= r.Count; i++) r[i].Apply();
        }, 2);

        public void CleanBoxes(Office.IRibbonControl c) => Run(() =>
        {
            var r = Selected();
            float h = r[1].Height;
            for (int i = 1; i <= r.Count; i++)
            {
                if (r[i].Type == Office.MsoShapeType.msoAutoShape)
                    r[i].AutoShapeType = Office.MsoAutoShapeType.msoShapeRectangle;
                r[i].Shadow.Visible = Office.MsoTriState.msoFalse;
                r[i].Line.Visible = Office.MsoTriState.msoFalse;
                if (i > 1) r[i].Height = h;
            }
        }, 1);
    }
}
