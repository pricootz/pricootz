using System;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Pricop.PowerPointTools
{
    internal static class PowerPointCommands
    {
        private static PowerPoint.ShapeRange Selected(int minimum = 1)
        {
            var app = Globals.ThisAddIn.Application;
            if (app == null || app.ActiveWindow == null) return null;
            var selection = app.ActiveWindow.Selection;
            if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes) return null;
            var range = selection.ShapeRange;
            return range.Count >= minimum ? range : null;
        }

        public static bool Execute(string command, bool showMessages = false)
        {
            try
            {
                switch ((command ?? "").Trim().ToLowerInvariant())
                {
                    case "align-left":
                        return WithSelection(2, showMessages, r => r.Align(Office.MsoAlignCmd.msoAlignLefts, Office.MsoTriState.msoFalse));
                    case "align-center":
                        return WithSelection(2, showMessages, r => r.Align(Office.MsoAlignCmd.msoAlignCenters, Office.MsoTriState.msoFalse));
                    case "align-right":
                        return WithSelection(2, showMessages, r => r.Align(Office.MsoAlignCmd.msoAlignRights, Office.MsoTriState.msoFalse));
                    case "align-top":
                        return WithSelection(2, showMessages, r => r.Align(Office.MsoAlignCmd.msoAlignTops, Office.MsoTriState.msoFalse));
                    case "align-middle":
                        return WithSelection(2, showMessages, r => r.Align(Office.MsoAlignCmd.msoAlignMiddles, Office.MsoTriState.msoFalse));
                    case "align-bottom":
                        return WithSelection(2, showMessages, r => r.Align(Office.MsoAlignCmd.msoAlignBottoms, Office.MsoTriState.msoFalse));

                    case "distribute-horizontal":
                        return WithSelection(3, showMessages, r => r.Distribute(Office.MsoDistributeCmd.msoDistributeHorizontally, Office.MsoTriState.msoFalse));
                    case "distribute-vertical":
                        return WithSelection(3, showMessages, r => r.Distribute(Office.MsoDistributeCmd.msoDistributeVertically, Office.MsoTriState.msoFalse));

                    case "same-width":
                        return WithSelection(2, showMessages, r =>
                        {
                            float v = r[1].Width;
                            for (int i = 2; i <= r.Count; i++) r[i].Width = v;
                        });
                    case "same-height":
                        return WithSelection(2, showMessages, r =>
                        {
                            float v = r[1].Height;
                            for (int i = 2; i <= r.Count; i++) r[i].Height = v;
                        });
                    case "same-size":
                        return WithSelection(2, showMessages, r =>
                        {
                            float w = r[1].Width, h = r[1].Height;
                            for (int i = 2; i <= r.Count; i++) { r[i].Width = w; r[i].Height = h; }
                        });

                    case "rectangle":
                        return WithSelection(1, showMessages, r =>
                        {
                            for (int i = 1; i <= r.Count; i++)
                                if (r[i].Type == Office.MsoShapeType.msoAutoShape)
                                    r[i].AutoShapeType = Office.MsoAutoShapeType.msoShapeRectangle;
                        });
                    case "rounded-rectangle":
                        return WithSelection(1, showMessages, r =>
                        {
                            for (int i = 1; i <= r.Count; i++)
                                if (r[i].Type == Office.MsoShapeType.msoAutoShape)
                                    r[i].AutoShapeType = Office.MsoAutoShapeType.msoShapeRoundedRectangle;
                        });

                    case "shadow-off":
                        return WithSelection(1, showMessages, r =>
                        {
                            for (int i = 1; i <= r.Count; i++) r[i].Shadow.Visible = Office.MsoTriState.msoFalse;
                        });
                    case "shadow-on":
                        return WithSelection(1, showMessages, r =>
                        {
                            for (int i = 1; i <= r.Count; i++)
                            {
                                var s = r[i].Shadow;
                                s.Visible = Office.MsoTriState.msoTrue;
                                s.Transparency = 0.55f;
                                s.Blur = 7f;
                                s.OffsetX = 1.5f;
                                s.OffsetY = 2f;
                            }
                        });

                    case "border-off":
                        return WithSelection(1, showMessages, r =>
                        {
                            for (int i = 1; i <= r.Count; i++) r[i].Line.Visible = Office.MsoTriState.msoFalse;
                        });
                    case "border-on":
                        return WithSelection(1, showMessages, r =>
                        {
                            for (int i = 1; i <= r.Count; i++)
                            {
                                r[i].Line.Visible = Office.MsoTriState.msoTrue;
                                r[i].Line.Weight = 1f;
                            }
                        });

                    case "match-style":
                        return WithSelection(2, showMessages, r =>
                        {
                            r[1].PickUp();
                            for (int i = 2; i <= r.Count; i++) r[i].Apply();
                        });

                    case "clean-boxes":
                        return WithSelection(1, showMessages, r =>
                        {
                            float h = r[1].Height;
                            for (int i = 1; i <= r.Count; i++)
                            {
                                if (r[i].Type == Office.MsoShapeType.msoAutoShape)
                                    r[i].AutoShapeType = Office.MsoAutoShapeType.msoShapeRectangle;
                                r[i].Shadow.Visible = Office.MsoTriState.msoFalse;
                                r[i].Line.Visible = Office.MsoTriState.msoFalse;
                                if (i > 1) r[i].Height = h;
                            }
                        });
                    default:
                        return false;
                }
            }
            catch (Exception ex)
            {
                if (showMessages)
                    MessageBox.Show(ex.Message, "Pricop Tools", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }

        private static bool WithSelection(int minimum, bool showMessages, Action<PowerPoint.ShapeRange> action)
        {
            var r = Selected(minimum);
            if (r == null)
            {
                if (showMessages)
                    MessageBox.Show(minimum <= 1 ? "Seleziona almeno una forma." : "Seleziona almeno " + minimum + " elementi.",
                        "Pricop Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            action(r);
            return true;
        }
    }
}
