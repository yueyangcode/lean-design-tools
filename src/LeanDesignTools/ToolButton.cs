using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace LeanDesignTools
{
    public enum IconKind
    {
        AlignLeft, AlignCenter, AlignRight, AlignTop, AlignMiddle, AlignBottom, DistributeH, DistributeV,
        SameWidth, SameHeight, SameSize, StretchH, StretchV,
        DockLeft, DockCenterH, DockRight, DockTop, DockCenterV, DockBottom, SlideCenter,
        SelectAll, Group, Ungroup, BringFront, Forward, Backward, SendBack, Eye,
        Union, Combine, Fragment, Intersect, Subtract,
        CopyPosition, PastePosition, CopyWidth, PasteWidth, CopyHeight, PasteHeight, CopySize, PasteSize,
        SnapLeft, SnapRight, SnapUp, SnapDown,
        RotateLeft, RotateRight, FlipH, FlipV,
        SplitParagraph, CombineParagraph, SplitLine, CombineLine, ClearText, ClearMargins, ResetMargins
    }

    internal sealed class ToolButton : Button
    {
        private bool _available = true;
        public IconKind Icon { get; set; }
        public int MinimumSelection { get; set; }
        public bool Available
        {
            get => _available;
            set { if (_available == value) return; _available = value; Invalidate(); }
        }

        public ToolButton()
        {
            FlatStyle = FlatStyle.Flat;
            BackColor = Color.White;
            Margin = new Padding(0);
            Size = new Size(51, 29);
            TabStop = false;
            Cursor = Cursors.Hand;
            FlatAppearance.BorderSize = 0;
            FlatAppearance.MouseOverBackColor = Color.FromArgb(247, 247, 247);
            FlatAppearance.MouseDownBackColor = Color.FromArgb(240, 240, 240);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var color = Available ? Color.FromArgb(244, 78, 72) : Color.FromArgb(165, 165, 165);
            using (var pen = new Pen(color, 1f))
            using (var brush = new SolidBrush(color))
            {
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;
                pen.LineJoin = LineJoin.Round;
                DrawIcon(e.Graphics, pen, brush, new Rectangle((Width - 17) / 2, (Height - 15) / 2, 17, 15));
            }
        }

        private void DrawIcon(Graphics g, Pen p, Brush b, Rectangle r)
        {
            switch (Icon)
            {
                case IconKind.AlignLeft: DrawAlignH(g, p, r, 0); break;
                case IconKind.AlignCenter: DrawAlignH(g, p, r, 1); break;
                case IconKind.AlignRight: DrawAlignH(g, p, r, 2); break;
                case IconKind.AlignTop: DrawAlignV(g, p, r, 0); break;
                case IconKind.AlignMiddle: DrawAlignV(g, p, r, 1); break;
                case IconKind.AlignBottom: DrawAlignV(g, p, r, 2); break;
                case IconKind.DistributeH: DrawDistribution(g, p, r, true); break;
                case IconKind.DistributeV: DrawDistribution(g, p, r, false); break;
                case IconKind.SameWidth: DrawSize(g, p, r, true, false); break;
                case IconKind.SameHeight: DrawSize(g, p, r, false, true); break;
                case IconKind.SameSize: DrawSize(g, p, r, true, true); break;
                case IconKind.StretchH: DrawStretch(g, p, r, true); break;
                case IconKind.StretchV: DrawStretch(g, p, r, false); break;
                case IconKind.DockLeft: DrawDock(g, p, b, r, -1, 2); break;
                case IconKind.DockCenterH: DrawDock(g, p, b, r, 0, 2); break;
                case IconKind.DockRight: DrawDock(g, p, b, r, 1, 2); break;
                case IconKind.DockTop: DrawDock(g, p, b, r, 2, -1); break;
                case IconKind.DockCenterV: DrawDock(g, p, b, r, 2, 0); break;
                case IconKind.DockBottom: DrawDock(g, p, b, r, 2, 1); break;
                case IconKind.SlideCenter: DrawSlideCenter(g, p, r); break;
                case IconKind.SelectAll: g.DrawRectangle(p, r.X + 1, r.Y + 1, 14, 12); g.DrawRectangle(p, r.X + 5, r.Y + 4, 6, 6); break;
                case IconKind.Group: DrawGroup(g, p, r, false); break;
                case IconKind.Ungroup: DrawGroup(g, p, r, true); break;
                case IconKind.BringFront: DrawLayers(g, p, b, r, 0); break;
                case IconKind.Forward: DrawLayers(g, p, b, r, 1); break;
                case IconKind.Backward: DrawLayers(g, p, b, r, 2); break;
                case IconKind.SendBack: DrawLayers(g, p, b, r, 3); break;
                case IconKind.Eye: g.DrawEllipse(p, r.X + 1, r.Y + 3, 14, 8); g.FillEllipse(b, r.X + 7, r.Y + 6, 2, 2); break;
                case IconKind.Union: DrawVector(g, p, b, r, 0); break;
                case IconKind.Combine: DrawVector(g, p, b, r, 1); break;
                case IconKind.Fragment: DrawVector(g, p, b, r, 2); break;
                case IconKind.Intersect: DrawVector(g, p, b, r, 3); break;
                case IconKind.Subtract: DrawVector(g, p, b, r, 4); break;
                case IconKind.CopyPosition: DrawTransfer(g, p, r, 0, false); break;
                case IconKind.PastePosition: DrawTransfer(g, p, r, 0, true); break;
                case IconKind.CopyWidth: DrawTransfer(g, p, r, 1, false); break;
                case IconKind.PasteWidth: DrawTransfer(g, p, r, 1, true); break;
                case IconKind.CopyHeight: DrawTransfer(g, p, r, 2, false); break;
                case IconKind.PasteHeight: DrawTransfer(g, p, r, 2, true); break;
                case IconKind.CopySize: DrawTransfer(g, p, r, 3, false); break;
                case IconKind.PasteSize: DrawTransfer(g, p, r, 3, true); break;
                case IconKind.SnapLeft: DrawSnap(g, p, r, 0); break;
                case IconKind.SnapRight: DrawSnap(g, p, r, 1); break;
                case IconKind.SnapUp: DrawSnap(g, p, r, 2); break;
                case IconKind.SnapDown: DrawSnap(g, p, r, 3); break;
                case IconKind.RotateLeft: DrawRotate(g, p, r, false); break;
                case IconKind.RotateRight: DrawRotate(g, p, r, true); break;
                case IconKind.FlipH: DrawFlip(g, p, r, true); break;
                case IconKind.FlipV: DrawFlip(g, p, r, false); break;
                case IconKind.SplitParagraph: DrawSplit(g, p, r, 0); break;
                case IconKind.CombineParagraph: DrawSplit(g, p, r, 1); break;
                case IconKind.SplitLine: DrawSplit(g, p, r, 2); break;
                case IconKind.CombineLine: DrawSplit(g, p, r, 3); break;
                case IconKind.ClearText: DrawClearText(g, p, r); break;
                case IconKind.ClearMargins: DrawTrash(g, p, r); break;
                case IconKind.ResetMargins: DrawReset(g, p, r); break;
            }
        }

        private static void DrawAlignH(Graphics g, Pen p, Rectangle r, int mode)
        {
            var x = mode == 0 ? r.Left + 2 : mode == 1 ? r.Left + r.Width / 2 : r.Right - 2;
            g.DrawLine(p, x, r.Top + 1, x, r.Bottom - 1);
            var widths = new[] { 11, 7 };
            for (var i = 0; i < 2; i++)
            {
                var y = r.Top + 4 + i * 6;
                var left = mode == 0 ? x : mode == 1 ? x - widths[i] / 2f : x - widths[i];
                g.DrawLine(p, left, y, left + widths[i], y);
                g.DrawLine(p, left, y + 1, left + widths[i], y + 1);
            }
        }

        private static void DrawAlignV(Graphics g, Pen p, Rectangle r, int mode)
        {
            var y = mode == 0 ? r.Top + 2 : mode == 1 ? r.Top + r.Height / 2 : r.Bottom - 2;
            g.DrawLine(p, r.Left + 1, y, r.Right - 1, y);
            var heights = new[] { 11, 7 };
            for (var i = 0; i < 2; i++)
            {
                var x = r.Left + 4 + i * 7;
                var top = mode == 0 ? y : mode == 1 ? y - heights[i] / 2f : y - heights[i];
                g.DrawLine(p, x, top, x, top + heights[i]);
                g.DrawLine(p, x + 1, top, x + 1, top + heights[i]);
            }
        }

        private static void DrawDistribution(Graphics g, Pen p, Rectangle r, bool horizontal)
        {
            if (horizontal)
            {
                g.DrawLine(p, r.Left + 1, r.Top + 1, r.Left + 1, r.Bottom - 1);
                g.DrawLine(p, r.Right - 1, r.Top + 1, r.Right - 1, r.Bottom - 1);
                for (var i = 0; i < 3; i++) g.DrawLine(p, r.Left + 4 + i * 4, r.Top + 4, r.Left + 4 + i * 4, r.Bottom - 4);
            }
            else
            {
                g.DrawLine(p, r.Left + 1, r.Top + 1, r.Right - 1, r.Top + 1);
                g.DrawLine(p, r.Left + 1, r.Bottom - 1, r.Right - 1, r.Bottom - 1);
                for (var i = 0; i < 3; i++) g.DrawLine(p, r.Left + 4, r.Top + 4 + i * 4, r.Right - 4, r.Top + 4 + i * 4);
            }
        }

        private static void DrawSize(Graphics g, Pen p, Rectangle r, bool width, bool height)
        {
            g.DrawRectangle(p, r.Left + 2, r.Top + 2, 13, 11);
            if (width && height) DrawDiagonalArrow(g, p, r.Left + 5, r.Top + 10, r.Right - 5, r.Top + 5);
            else if (width) DrawDoubleArrowH(g, p, r.Left + 4, r.Right - 4, r.Top + 7);
            else DrawDoubleArrowV(g, p, r.Left + 8, r.Top + 4, r.Bottom - 4);
        }

        private static void DrawStretch(Graphics g, Pen p, Rectangle r, bool horizontal)
        {
            var oldDash = p.DashStyle;
            p.DashStyle = DashStyle.Dot;
            g.DrawRectangle(p, r.Left + 1, r.Top + 1, 14, 12);
            p.DashStyle = oldDash;
            if (horizontal) DrawDoubleArrowH(g, p, r.Left + 3, r.Right - 3, r.Top + 7);
            else DrawDoubleArrowV(g, p, r.Left + 8, r.Top + 3, r.Bottom - 3);
        }

        private static void DrawDock(Graphics g, Pen p, Brush b, Rectangle r, int horizontal, int vertical)
        {
            if (horizontal != 2)
            {
                var guide = horizontal == -1 ? r.Left + 2 : horizontal == 0 ? r.Left + 8 : r.Right - 2;
                g.DrawLine(p, guide, r.Top + 1, guide, r.Bottom - 1);
                if (horizontal == -1) DrawArrowH(g, p, r.Right - 2, guide + 2, r.Top + 7);
                else if (horizontal == 1) DrawArrowH(g, p, r.Left + 2, guide - 2, r.Top + 7);
                else { DrawArrowH(g, p, r.Left + 1, guide - 2, r.Top + 7); DrawArrowH(g, p, r.Right - 1, guide + 2, r.Top + 7); }
            }
            else
            {
                var guide = vertical == -1 ? r.Top + 2 : vertical == 0 ? r.Top + 7 : r.Bottom - 2;
                g.DrawLine(p, r.Left + 1, guide, r.Right - 1, guide);
                if (vertical == -1) DrawArrowV(g, p, r.Left + 8, r.Bottom - 1, guide + 2);
                else if (vertical == 1) DrawArrowV(g, p, r.Left + 8, r.Top + 1, guide - 2);
                else { DrawArrowV(g, p, r.Left + 8, r.Top + 1, guide - 2); DrawArrowV(g, p, r.Left + 8, r.Bottom - 1, guide + 2); }
            }
        }

        private static void DrawSlideCenter(Graphics g, Pen p, Rectangle r)
        {
            var oldDash = p.DashStyle;
            p.DashStyle = DashStyle.Dot;
            g.DrawRectangle(p, r.Left + 1, r.Top + 1, 14, 12);
            p.DashStyle = oldDash;
            g.DrawLine(p, r.Left + 8, r.Top + 4, r.Left + 8, r.Bottom - 4);
            g.DrawLine(p, r.Left + 5, r.Top + 7, r.Right - 5, r.Top + 7);
        }

        private static void DrawDoubleArrowH(Graphics g, Pen p, int left, int right, int y)
        {
            g.DrawLine(p, left, y, right, y);
            g.DrawLine(p, left, y, left + 3, y - 2); g.DrawLine(p, left, y, left + 3, y + 2);
            g.DrawLine(p, right, y, right - 3, y - 2); g.DrawLine(p, right, y, right - 3, y + 2);
        }

        private static void DrawDoubleArrowV(Graphics g, Pen p, int x, int top, int bottom)
        {
            g.DrawLine(p, x, top, x, bottom);
            g.DrawLine(p, x, top, x - 2, top + 3); g.DrawLine(p, x, top, x + 2, top + 3);
            g.DrawLine(p, x, bottom, x - 2, bottom - 3); g.DrawLine(p, x, bottom, x + 2, bottom - 3);
        }

        private static void DrawDiagonalArrow(Graphics g, Pen p, int x1, int y1, int x2, int y2)
        {
            g.DrawLine(p, x1, y1, x2, y2);
            g.DrawLine(p, x1, y1, x1 + 1, y1 - 3); g.DrawLine(p, x1, y1, x1 + 3, y1 - 1);
            g.DrawLine(p, x2, y2, x2 - 1, y2 + 3); g.DrawLine(p, x2, y2, x2 - 3, y2 + 1);
        }

        private static void DrawArrowH(Graphics g, Pen p, int from, int to, int y)
        {
            g.DrawLine(p, from, y, to, y);
            var d = from < to ? -1 : 1;
            g.DrawLine(p, to, y, to + d * 3, y - 2); g.DrawLine(p, to, y, to + d * 3, y + 2);
        }

        private static void DrawArrowV(Graphics g, Pen p, int x, int from, int to)
        {
            g.DrawLine(p, x, from, x, to);
            var d = from < to ? -1 : 1;
            g.DrawLine(p, x, to, x - 2, to + d * 3); g.DrawLine(p, x, to, x + 2, to + d * 3);
        }

        private static void DrawGroup(Graphics g, Pen p, Rectangle r, bool split)
        {
            g.DrawRectangle(p, r.Left + 1, r.Top + 2, 8, 8); g.DrawRectangle(p, r.Left + 7, r.Top + 5, 8, 8);
            if (split) g.DrawLine(p, r.Left + 2, r.Bottom - 1, r.Right - 2, r.Top + 1);
        }

        private static void DrawLayers(Graphics g, Pen p, Brush b, Rectangle r, int mode)
        {
            g.DrawRectangle(p, r.Left + 2, r.Top + 5, 9, 8); g.DrawRectangle(p, r.Left + 6, r.Top + 1, 9, 8);
            if (mode < 2) g.FillRectangle(b, r.Left + 7, r.Top + 2, 7, 6); else g.FillRectangle(b, r.Left + 3, r.Top + 6, 7, 6);
        }

        private static void DrawVector(Graphics g, Pen p, Brush b, Rectangle r, int mode)
        {
            g.DrawEllipse(p, r.Left + 1, r.Top + 3, 9, 9); g.DrawRectangle(p, r.Left + 7, r.Top + 1, 8, 9);
            if (mode == 0) { g.FillEllipse(b, r.Left + 2, r.Top + 4, 8, 8); g.FillRectangle(b, r.Left + 8, r.Top + 2, 7, 8); }
            if (mode == 3) g.FillRectangle(b, r.Left + 8, r.Top + 4, 3, 5);
            if (mode == 4) g.DrawLine(p, r.Left + 5, r.Top + 7, r.Right - 1, r.Top + 7);
            if (mode == 2) { g.DrawLine(p, r.Left + 8, r.Top + 1, r.Left + 8, r.Bottom - 1); g.DrawLine(p, r.Left + 1, r.Top + 8, r.Right - 1, r.Top + 8); }
        }

        private static void DrawTransfer(Graphics g, Pen p, Rectangle r, int part, bool paste)
        {
            g.DrawRectangle(p, r.Left + 5, r.Top + 2, 8, 10);
            var arrowY = r.Bottom - 1;
            if (paste)
            {
                g.DrawLine(p, r.Left + 1, arrowY, r.Left + 7, arrowY);
                g.DrawLine(p, r.Left + 5, arrowY - 2, r.Left + 7, arrowY);
            }
            else
            {
                g.DrawLine(p, r.Right - 7, arrowY, r.Right - 1, arrowY);
                g.DrawLine(p, r.Right - 3, arrowY - 2, r.Right - 1, arrowY);
            }
            if (part == 0)
            {
                g.DrawLine(p, r.Left + 7, r.Top + 7, r.Left + 11, r.Top + 7);
                g.DrawLine(p, r.Left + 9, r.Top + 5, r.Left + 9, r.Top + 9);
            }
            else if (part == 1)
            {
                g.DrawLine(p, r.Left + 6, r.Top + 7, r.Left + 12, r.Top + 7);
                g.DrawLine(p, r.Left + 6, r.Top + 5, r.Left + 6, r.Top + 9);
                g.DrawLine(p, r.Left + 12, r.Top + 5, r.Left + 12, r.Top + 9);
            }
            else if (part == 2)
            {
                g.DrawLine(p, r.Left + 9, r.Top + 3, r.Left + 9, r.Top + 11);
                g.DrawLine(p, r.Left + 7, r.Top + 3, r.Left + 11, r.Top + 3);
                g.DrawLine(p, r.Left + 7, r.Top + 11, r.Left + 11, r.Top + 11);
            }
            else
            {
                g.DrawLine(p, r.Left + 6, r.Top + 10, r.Left + 12, r.Top + 4);
                g.DrawLine(p, r.Left + 9, r.Top + 4, r.Left + 12, r.Top + 4);
                g.DrawLine(p, r.Left + 12, r.Top + 4, r.Left + 12, r.Top + 7);
            }
        }

        private static void DrawRotate(Graphics g, Pen p, Rectangle r, bool right)
        {
            g.DrawArc(p, r.Left + 3, r.Top + 2, 10, 10, right ? 210 : -30, 250);
            var x = right ? r.Right - 2 : r.Left + 2; g.DrawLine(p, x, r.Top + 8, x + (right ? -4 : 4), r.Top + 8); g.DrawLine(p, x, r.Top + 8, x, r.Top + 4);
        }

        private static void DrawFlip(Graphics g, Pen p, Rectangle r, bool horizontal)
        {
            if (horizontal) { g.DrawLine(p, r.Left + 8, r.Top, r.Left + 8, r.Bottom); g.DrawRectangle(p, r.Left + 1, r.Top + 4, 5, 7); g.DrawRectangle(p, r.Left + 10, r.Top + 4, 5, 7); }
            else { g.DrawLine(p, r.Left, r.Top + 7, r.Right, r.Top + 7); g.DrawRectangle(p, r.Left + 5, r.Top, 7, 5); g.DrawRectangle(p, r.Left + 5, r.Top + 9, 7, 5); }
        }

        private static void DrawSnap(Graphics g, Pen p, Rectangle r, int direction)
        {
            g.DrawRectangle(p, r.Left + 6, r.Top + 4, 6, 6);
            if (direction < 2)
            {
                var x = direction == 0 ? r.Left + 1 : r.Right - 1;
                g.DrawLine(p, x, r.Top + 2, x, r.Bottom - 2);
                g.DrawLine(p, direction == 0 ? r.Left + 5 : r.Right - 5, r.Top + 7, x, r.Top + 7);
            }
            else
            {
                var y = direction == 2 ? r.Top + 1 : r.Bottom - 1;
                g.DrawLine(p, r.Left + 2, y, r.Right - 2, y);
                g.DrawLine(p, r.Left + 9, direction == 2 ? r.Top + 4 : r.Bottom - 4, r.Left + 9, y);
            }
        }

        private static void DrawSplit(Graphics g, Pen p, Rectangle r, int mode)
        {
            using (var font = new Font("Segoe UI", 7f))
            using (var brush = new SolidBrush(p.Color))
            {
                g.DrawString("A", font, brush, r.Left, r.Top + 1);
                g.DrawString("A", font, brush, r.Left + 10, r.Top + 1);
            }
            if (mode == 0 || mode == 2) g.DrawLine(p, r.Left + 8, r.Top + 1, r.Left + 8, r.Bottom - 1);
            else g.DrawLine(p, r.Left + 6, r.Top + 7, r.Left + 10, r.Top + 7);
            if (mode >= 2) g.DrawLine(p, r.Left + 1, r.Bottom - 1, r.Right - 1, r.Bottom - 1);
        }

        private static void DrawTrash(Graphics g, Pen p, Rectangle r)
        {
            g.DrawRectangle(p, r.Left + 4, r.Top + 5, 8, 8);
            g.DrawLine(p, r.Left + 3, r.Top + 3, r.Left + 13, r.Top + 3);
            g.DrawLine(p, r.Left + 6, r.Top + 1, r.Left + 10, r.Top + 1);
        }

        private static void DrawClearText(Graphics g, Pen p, Rectangle r)
        {
            using (var font = new Font("Segoe UI", 7f))
            using (var brush = new SolidBrush(p.Color))
                g.DrawString("A", font, brush, r.Left, r.Top + 1);
            g.DrawPolygon(p, new[]
            {
                new Point(r.Right - 7, r.Top + 3), new Point(r.Right - 2, r.Top + 8),
                new Point(r.Right - 7, r.Top + 13), new Point(r.Right - 12, r.Top + 8)
            });
        }

        private static void DrawReset(Graphics g, Pen p, Rectangle r)
        {
            g.DrawArc(p, r.Left + 2, r.Top + 2, 12, 11, 35, 290);
            g.DrawLine(p, r.Left + 2, r.Top + 7, r.Left + 2, r.Top + 2);
            g.DrawLine(p, r.Left + 2, r.Top + 2, r.Left + 7, r.Top + 2);
        }
    }
}
