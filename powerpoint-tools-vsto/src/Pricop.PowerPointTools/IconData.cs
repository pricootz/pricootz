using System;
using System.Drawing;
using System.Windows.Forms;

namespace Pricop.PowerPointTools
{
    internal static class IconData
    {
        private static readonly Color Accent = Color.FromArgb(19, 183, 166);
        private static readonly Pen WhitePen = new Pen(Color.White, 2f);
        private static readonly Brush WhiteBrush = Brushes.White;

        public static object GetPicture(string key)
        {
            using (var bmp = Draw(key, 32))
                return PictureDispHost.ToPictureDisp(bmp);
        }

        private static Bitmap Draw(string key, int size)
        {
            var bmp = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using (var bg = new SolidBrush(Accent))
                    g.FillEllipse(bg, 2, 2, size - 4, size - 4);

                switch (key)
                {
                    case "left":
                        g.DrawLine(WhitePen, 8, 8, 8, 24);
                        g.DrawLine(WhitePen, 11, 11, 24, 11);
                        g.DrawLine(WhitePen, 11, 16, 20, 16);
                        g.DrawLine(WhitePen, 11, 21, 24, 21);
                        break;
                    case "center":
                        g.DrawLine(Pens.White, 16, 7, 16, 25);
                        g.DrawLine(WhitePen, 8, 11, 24, 11);
                        g.DrawLine(WhitePen, 10, 16, 22, 16);
                        g.DrawLine(WhitePen, 7, 21, 25, 21);
                        break;
                    case "right":
                        g.DrawLine(WhitePen, 24, 8, 24, 24);
                        g.DrawLine(WhitePen, 8, 11, 21, 11);
                        g.DrawLine(WhitePen, 12, 16, 21, 16);
                        g.DrawLine(WhitePen, 8, 21, 21, 21);
                        break;
                    case "top":
                        g.DrawLine(WhitePen, 8, 8, 24, 8);
                        g.DrawLine(WhitePen, 11, 11, 11, 23);
                        g.DrawLine(WhitePen, 16, 11, 16, 20);
                        g.DrawLine(WhitePen, 21, 11, 21, 24);
                        break;
                    case "middle":
                        g.DrawLine(Pens.White, 7, 16, 25, 16);
                        g.DrawLine(WhitePen, 11, 8, 11, 24);
                        g.DrawLine(WhitePen, 16, 10, 16, 22);
                        g.DrawLine(WhitePen, 21, 7, 21, 25);
                        break;
                    case "bottom":
                        g.DrawLine(WhitePen, 8, 24, 24, 24);
                        g.DrawLine(WhitePen, 11, 9, 11, 21);
                        g.DrawLine(WhitePen, 16, 12, 16, 21);
                        g.DrawLine(WhitePen, 21, 8, 21, 21);
                        break;
                    case "dist_h":
                        g.DrawLine(Pens.White, 6, 9, 6, 23);
                        g.DrawLine(Pens.White, 26, 9, 26, 23);
                        g.FillRectangle(WhiteBrush, 10, 12, 3, 8);
                        g.FillRectangle(WhiteBrush, 15, 12, 3, 8);
                        g.FillRectangle(WhiteBrush, 20, 12, 3, 8);
                        break;
                    case "dist_v":
                        g.DrawLine(Pens.White, 9, 6, 23, 6);
                        g.DrawLine(Pens.White, 9, 26, 23, 26);
                        g.FillRectangle(WhiteBrush, 12, 10, 8, 3);
                        g.FillRectangle(WhiteBrush, 12, 15, 8, 3);
                        g.FillRectangle(WhiteBrush, 12, 20, 8, 3);
                        break;
                    case "width":
                        g.DrawRectangle(WhitePen, 9, 10, 14, 12);
                        g.DrawLine(WhitePen, 6, 16, 9, 16);
                        g.DrawLine(WhitePen, 23, 16, 26, 16);
                        break;
                    case "height":
                        g.DrawRectangle(WhitePen, 10, 9, 12, 14);
                        g.DrawLine(WhitePen, 16, 6, 16, 9);
                        g.DrawLine(WhitePen, 16, 23, 16, 26);
                        break;
                    case "size":
                        g.DrawRectangle(WhitePen, 9, 9, 14, 14);
                        g.DrawLine(WhitePen, 7, 7, 12, 7);
                        g.DrawLine(WhitePen, 7, 7, 7, 12);
                        g.DrawLine(WhitePen, 25, 25, 20, 25);
                        g.DrawLine(WhitePen, 25, 25, 25, 20);
                        break;
                    case "rect":
                        g.DrawRectangle(WhitePen, 8, 10, 16, 12);
                        break;
                    case "round":
                        using (var p = RoundedRect(new Rectangle(8, 10, 16, 12), 4))
                            g.DrawPath(WhitePen, p);
                        break;
                    case "shadow_off":
                        g.DrawRectangle(WhitePen, 8, 8, 13, 11);
                        g.DrawLine(WhitePen, 8, 24, 24, 8);
                        break;
                    case "shadow_on":
                        using (var shadow = new SolidBrush(Color.FromArgb(100, 255, 255, 255)))
                            g.FillRectangle(shadow, 12, 13, 12, 10);
                        g.DrawRectangle(WhitePen, 8, 8, 13, 11);
                        break;
                    case "line_off":
                        g.DrawRectangle(WhitePen, 8, 9, 15, 13);
                        g.DrawLine(WhitePen, 7, 24, 25, 7);
                        break;
                    case "line_on":
                        using (var p = new Pen(Color.White, 3f)) g.DrawRectangle(p, 8, 9, 15, 13);
                        break;
                    case "copy":
                        g.DrawRectangle(WhitePen, 7, 8, 11, 12);
                        g.DrawRectangle(WhitePen, 13, 12, 12, 12);
                        break;
                    case "clean":
                        g.DrawRectangle(WhitePen, 8, 10, 15, 12);
                        g.DrawLine(WhitePen, 21, 7, 25, 11);
                        g.DrawLine(WhitePen, 23, 6, 23, 12);
                        break;
                }
            }
            return bmp;
        }

        private static System.Drawing.Drawing2D.GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            var p = new System.Drawing.Drawing2D.GraphicsPath();
            int d = radius * 2;
            p.AddArc(r.Left, r.Top, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Top, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.Left, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        private sealed class PictureDispHost : AxHost
        {
            private PictureDispHost() : base("") { }
            public static object ToPictureDisp(Image image) => GetIPictureDispFromPicture(image);
        }
    }
}
