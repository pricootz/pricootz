using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace Pricop.PowerPointTools
{
    internal static class IconData
    {
        private static readonly Color Accent = Color.FromArgb(32, 224, 194);
        private static readonly Color AccentMuted = Color.FromArgb(120, 32, 224, 194);
        private static readonly Color Ink = Color.FromArgb(248, 250, 252);
        private static readonly Color Secondary = Color.FromArgb(54, 86, 102);

        public static object GetPicture(string key)
        {
            using (var bmp = Draw(key, 32))
                return PictureDispHost.ToPictureDisp(bmp);
        }

        private static Bitmap Draw(string key, int size)
        {
            var bmp = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bmp))
            using (var whitePen = RoundedPen(Ink, 2.35f))
            using (var accentPen = RoundedPen(Accent, 2.35f))
            using (var thinWhite = RoundedPen(Ink, 1.8f))
            using (var whiteBrush = new SolidBrush(Ink))
            using (var accentBrush = new SolidBrush(Accent))
            using (var accentSoft = new SolidBrush(AccentMuted))
            using (var secondaryBrush = new SolidBrush(Secondary))
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;

                switch (key)
                {
                    case "left":
                        g.DrawLine(whitePen, 7, 6, 7, 26);
                        FillRound(g, accentBrush, 11, 9, 10, 6, 2);
                        FillRound(g, whiteBrush, 11, 18, 15, 6, 2);
                        break;

                    case "center":
                        g.DrawLine(thinWhite, 16, 5, 16, 27);
                        FillRound(g, accentBrush, 9, 9, 14, 6, 2);
                        FillRound(g, whiteBrush, 6, 18, 20, 6, 2);
                        break;

                    case "right":
                        g.DrawLine(whitePen, 25, 6, 25, 26);
                        FillRound(g, accentBrush, 11, 9, 10, 6, 2);
                        FillRound(g, whiteBrush, 6, 18, 15, 6, 2);
                        break;

                    case "top":
                        g.DrawLine(whitePen, 6, 7, 26, 7);
                        FillRound(g, accentBrush, 9, 11, 6, 14, 2);
                        FillRound(g, secondaryBrush, 18, 11, 6, 10, 2);
                        break;

                    case "middle":
                        g.DrawLine(thinWhite, 5, 16, 27, 16);
                        FillRound(g, secondaryBrush, 8, 9, 5, 14, 2);
                        FillRound(g, accentBrush, 14, 6, 5, 20, 2);
                        FillRound(g, whiteBrush, 20, 10, 5, 12, 2);
                        break;

                    case "bottom":
                        g.DrawLine(whitePen, 6, 25, 26, 25);
                        FillRound(g, accentBrush, 9, 7, 6, 14, 2);
                        FillRound(g, whiteBrush, 18, 11, 6, 10, 2);
                        break;

                    case "dist_h":
                        g.DrawLine(whitePen, 5, 7, 5, 25);
                        g.DrawLine(whitePen, 27, 7, 27, 25);
                        FillRound(g, accentSoft, 9, 11, 4, 10, 1.5f);
                        FillRound(g, accentBrush, 14, 9, 4, 14, 1.5f);
                        FillRound(g, accentSoft, 19, 11, 4, 10, 1.5f);
                        break;

                    case "dist_v":
                        g.DrawLine(whitePen, 7, 5, 25, 5);
                        g.DrawLine(whitePen, 7, 27, 25, 27);
                        FillRound(g, accentSoft, 11, 9, 10, 4, 1.5f);
                        FillRound(g, accentBrush, 9, 14, 14, 4, 1.5f);
                        FillRound(g, accentSoft, 11, 19, 10, 4, 1.5f);
                        break;

                    case "width":
                        DrawRound(g, whitePen, 10, 10, 12, 12, 2.5f);
                        g.DrawLine(accentPen, 6, 8, 6, 24);
                        g.DrawLine(accentPen, 26, 8, 26, 24);
                        g.DrawLine(accentPen, 6, 8, 9, 8);
                        g.DrawLine(accentPen, 23, 8, 26, 8);
                        g.DrawLine(accentPen, 6, 24, 9, 24);
                        g.DrawLine(accentPen, 23, 24, 26, 24);
                        break;

                    case "height":
                        DrawRound(g, whitePen, 10, 10, 12, 12, 2.5f);
                        g.DrawLine(accentPen, 8, 6, 24, 6);
                        g.DrawLine(accentPen, 8, 26, 24, 26);
                        g.DrawLine(accentPen, 8, 6, 8, 9);
                        g.DrawLine(accentPen, 24, 6, 24, 9);
                        g.DrawLine(accentPen, 8, 23, 8, 26);
                        g.DrawLine(accentPen, 24, 23, 24, 26);
                        break;

                    case "size":
                        DrawRound(g, whitePen, 10, 10, 12, 12, 2.5f);
                        g.DrawLine(accentPen, 6, 11, 6, 6);
                        g.DrawLine(accentPen, 6, 6, 11, 6);
                        g.DrawLine(accentPen, 26, 11, 26, 6);
                        g.DrawLine(accentPen, 26, 6, 21, 6);
                        g.DrawLine(accentPen, 6, 21, 6, 26);
                        g.DrawLine(accentPen, 6, 26, 11, 26);
                        g.DrawLine(accentPen, 26, 21, 26, 26);
                        g.DrawLine(accentPen, 26, 26, 21, 26);
                        break;

                    case "rect":
                        using (var fill = new SolidBrush(Color.FromArgb(45, Accent)))
                            g.FillRectangle(fill, 7, 10, 18, 12);
                        DrawRound(g, whitePen, 7, 10, 18, 12, 2);
                        g.DrawLine(accentPen, 10, 21, 22, 11);
                        break;

                    case "round":
                        using (var fill = new SolidBrush(Color.FromArgb(45, Accent)))
                        using (var p = RoundedRect(new RectangleF(7, 9, 18, 14), 5))
                        {
                            g.FillPath(fill, p);
                            g.DrawPath(whitePen, p);
                        }
                        g.DrawArc(accentPen, 10, 12, 12, 8, 18, 145);
                        break;

                    case "shadow_off":
                        FillRound(g, accentSoft, 12, 12, 12, 11, 2.5f);
                        DrawRound(g, whitePen, 7, 7, 14, 12, 2.5f);
                        g.DrawLine(whitePen, 6, 26, 26, 6);
                        break;

                    case "shadow_on":
                        FillRound(g, accentBrush, 13, 13, 12, 11, 2.5f);
                        DrawRound(g, whitePen, 7, 7, 15, 13, 2.5f);
                        break;

                    case "line_off":
                        DrawDashedRound(g, whitePen, 7, 7, 18, 18, 3.5f);
                        g.DrawLine(whitePen, 6, 26, 26, 6);
                        break;

                    case "line_on":
                        using (var p = RoundedPen(Ink, 3f))
                            DrawRound(g, p, 7, 7, 18, 18, 3.5f);
                        g.DrawLine(accentPen, 9, 9, 23, 9);
                        break;

                    case "copy":
                        using (var p = RoundedPen(Ink, 2.1f))
                        {
                            g.DrawLine(p, 18, 8, 24, 14);
                            g.DrawLine(p, 11, 21, 20, 12);
                        }
                        FillRound(g, accentBrush, 18, 7, 6, 6, 2);
                        g.DrawArc(whitePen, 7, 15, 12, 10, 120, 190);
                        g.DrawLine(accentPen, 10, 24, 15, 20);
                        break;

                    case "clean":
                        DrawSpark(g, accentBrush, 16, 9, 6);
                        DrawSpark(g, whiteBrush, 9, 19, 4.5f);
                        DrawSpark(g, whiteBrush, 23, 21, 3.5f);
                        break;

                    default:
                        DrawRound(g, whitePen, 7, 7, 18, 18, 4);
                        break;
                }
            }
            return bmp;
        }

        private static Pen RoundedPen(Color color, float width)
        {
            var p = new Pen(color, width);
            p.StartCap = LineCap.Round;
            p.EndCap = LineCap.Round;
            p.LineJoin = LineJoin.Round;
            return p;
        }

        private static void FillRound(Graphics g, Brush brush, float x, float y, float w, float h, float radius)
        {
            using (var p = RoundedRect(new RectangleF(x, y, w, h), radius))
                g.FillPath(brush, p);
        }

        private static void DrawRound(Graphics g, Pen pen, float x, float y, float w, float h, float radius)
        {
            using (var p = RoundedRect(new RectangleF(x, y, w, h), radius))
                g.DrawPath(pen, p);
        }

        private static void DrawDashedRound(Graphics g, Pen source, float x, float y, float w, float h, float radius)
        {
            using (var p = RoundedPen(source.Color, source.Width))
            {
                p.DashStyle = DashStyle.Dash;
                DrawRound(g, p, x, y, w, h, radius);
            }
        }

        private static void DrawSpark(Graphics g, Brush brush, float cx, float cy, float r)
        {
            var pts = new[]
            {
                new PointF(cx, cy-r),
                new PointF(cx+r*0.34f, cy-r*0.34f),
                new PointF(cx+r, cy),
                new PointF(cx+r*0.34f, cy+r*0.34f),
                new PointF(cx, cy+r),
                new PointF(cx-r*0.34f, cy+r*0.34f),
                new PointF(cx-r, cy),
                new PointF(cx-r*0.34f, cy-r*0.34f)
            };
            g.FillPolygon(brush, pts);
        }

        private static GraphicsPath RoundedRect(RectangleF r, float radius)
        {
            var p = new GraphicsPath();
            float d = radius * 2f;
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
