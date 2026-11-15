using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace LeanDesignTools
{
    public enum AlignMode { Left, Center, Right, Top, Middle, Bottom }

    public static class ShapeOperations
    {
        private static Box _copied;
        private static bool _hasCopiedPosition;
        private static bool _hasCopiedWidth;
        private static bool _hasCopiedHeight;

        public static void Run(Action<dynamic, dynamic> action, int minimum = 1)
        {
            try
            {
                dynamic app = DesignPane.PowerPoint;
                if (app == null) throw new InvalidOperationException("PowerPoint 尚未连接。");
                dynamic selection = app.ActiveWindow.Selection;
                var selectionType = (int)selection.Type;
                if (minimum > 0 && ((selectionType != 2 && selectionType != 3) || (int)selection.ShapeRange.Count < minimum))
                    throw new InvalidOperationException($"请至少选择 {minimum} 个形状。");
                app.StartNewUndoEntry();
                action(app, minimum > 0 ? selection.ShapeRange : null);
            }
            catch (Exception ex)
            {
                var message = ex is InvalidOperationException ? ex.Message : "操作失败：" + ex.Message;
                MessageBox.Show(message, "设计工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public static void Align(AlignMode mode) => Run((app, range) =>
        {
            dynamic anchor = range.Item(1);
            var anchorBox = Read(anchor, 1);
            for (var i = 2; i <= (int)range.Count; i++)
            {
                dynamic item = range.Item(i);
                var box = Read(item, i);
                switch (mode)
                {
                    case AlignMode.Left: item.Left = Geometry.AlignLeft(anchorBox); break;
                    case AlignMode.Center: item.Left = Geometry.AlignCenter(anchorBox, box); break;
                    case AlignMode.Right: item.Left = Geometry.AlignRight(anchorBox, box); break;
                    case AlignMode.Top: item.Top = Geometry.AlignTop(anchorBox); break;
                    case AlignMode.Middle: item.Top = Geometry.AlignMiddle(anchorBox, box); break;
                    case AlignMode.Bottom: item.Top = Geometry.AlignBottom(anchorBox, box); break;
                }
            }
        }, 2);

        public static void Distribute(bool horizontal) => Run((app, range) =>
        {
            var shapes = new Dictionary<int, dynamic>();
            var boxes = new List<Box>();
            for (var i = 1; i <= (int)range.Count; i++)
            {
                dynamic shape = range.Item(i);
                shapes[i] = shape;
                boxes.Add(Read(shape, i));
            }
            if (horizontal) Geometry.DistributeHorizontally(boxes); else Geometry.DistributeVertically(boxes);
            foreach (var box in boxes)
                if (horizontal) shapes[box.Id].Left = box.Left; else shapes[box.Id].Top = box.Top;
        }, 3);

        public static void MatchSize(bool width, bool height) => Run((app, range) =>
        {
            dynamic anchor = range.Item(1);
            for (var i = 2; i <= (int)range.Count; i++)
            {
                dynamic item = range.Item(i);
                if (width) item.Width = anchor.Width;
                if (height) item.Height = anchor.Height;
            }
        }, 2);

        public static void SetDimensions(decimal? widthCm, decimal? heightCm)
        {
            const float pointsPerCentimeter = 28.3464567f;
            Run((app, range) =>
            {
                for (var i = 1; i <= (int)range.Count; i++)
                {
                    dynamic shape = range.Item(i);
                    if (widthCm.HasValue) shape.Width = (float)widthCm.Value * pointsPerCentimeter;
                    if (heightCm.HasValue) shape.Height = (float)heightCm.Value * pointsPerCentimeter;
                }
            });
        }

        public static void Dock(int horizontal, int vertical) => Run((app, range) =>
        {
            float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
            float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;
            for (var i = 1; i <= (int)range.Count; i++)
            {
                dynamic shape = range.Item(i);
                if (horizontal == -1) shape.Left = 0f;
                if (horizontal == 0) shape.Left = (slideWidth - shape.Width) / 2f;
                if (horizontal == 1) shape.Left = slideWidth - shape.Width;
                if (vertical == -1) shape.Top = 0f;
                if (vertical == 0) shape.Top = (slideHeight - shape.Height) / 2f;
                if (vertical == 1) shape.Top = slideHeight - shape.Height;
            }
        });

        public static void Stretch(bool horizontal) => Run((app, range) =>
        {
            float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
            float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;
            for (var i = 1; i <= (int)range.Count; i++)
            {
                dynamic shape = range.Item(i);
                if (horizontal) { shape.Left = 0f; shape.Width = slideWidth; }
                else { shape.Top = 0f; shape.Height = slideHeight; }
            }
        });

        public static void Rotate(float degrees) => Run((app, range) =>
        {
            for (var i = 1; i <= (int)range.Count; i++) range.Item(i).Rotation += degrees;
        });

        public static void Flip(bool horizontal) => Run((app, range) =>
        {
            for (var i = 1; i <= (int)range.Count; i++) range.Item(i).Flip(horizontal ? 0 : 1);
        });

        public static void ZOrder(int command) => Run((app, range) =>
        {
            for (var i = 1; i <= (int)range.Count; i++) range.Item(i).ZOrder(command);
        });

        public static void Group() => Run((app, range) => range.Group(), 2);
        public static void Ungroup() => Run((app, range) => range.Ungroup());
        public static void ExecuteMso(string id, int minimum = 1) => Run((app, range) => app.CommandBars.ExecuteMso(id), minimum);

        public static void SelectAll()
        {
            try
            {
                dynamic app = DesignPane.PowerPoint;
                dynamic slide = app.ActiveWindow.View.Slide;
                if ((int)slide.Shapes.Count > 0) slide.Shapes.Range(Type.Missing).Select();
            }
            catch (Exception ex) { MessageBox.Show("选择失败：" + ex.Message, "设计工具"); }
        }

        public static void CopyGeometry(bool position, bool width, bool height) => Run((app, range) =>
        {
            var source = Read(range.Item(1), 1);
            if (_copied == null) _copied = new Box();
            if (position) { _copied.Left = source.Left; _copied.Top = source.Top; _hasCopiedPosition = true; }
            if (width) { _copied.Width = source.Width; _hasCopiedWidth = true; }
            if (height) { _copied.Height = source.Height; _hasCopiedHeight = true; }
        });

        public static void PasteGeometry(bool position, bool width, bool height) => Run((app, range) =>
        {
            if (position && !_hasCopiedPosition) throw new InvalidOperationException("请先复制位置。");
            if (width && !_hasCopiedWidth) throw new InvalidOperationException("请先复制宽度或大小。");
            if (height && !_hasCopiedHeight) throw new InvalidOperationException("请先复制高度或大小。");
            for (var i = 1; i <= (int)range.Count; i++)
            {
                dynamic shape = range.Item(i);
                if (position) { shape.Left = _copied.Left; shape.Top = _copied.Top; }
                if (width) shape.Width = _copied.Width;
                if (height) shape.Height = _copied.Height;
            }
        });

        public static void ClearText() => Run((app, range) =>
        {
            for (var i = 1; i <= (int)range.Count; i++)
            {
                dynamic shape = range.Item(i);
                try { if ((int)shape.HasTextFrame != 0) shape.TextFrame.TextRange.Text = ""; } catch { }
            }
        });

        public static void AutoSizeText() => Run((app, range) =>
        {
            for (var i = 1; i <= (int)range.Count; i++)
            {
                dynamic shape = range.Item(i);
                try { if ((int)shape.HasTextFrame != 0) shape.TextFrame.AutoSize = 1; } catch { }
            }
        });

        public static void Snap(SnapDirection direction) => Run((app, range) =>
        {
            dynamic slide = app.ActiveWindow.View.Slide;
            float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
            float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;
            var selectedIds = new HashSet<int>();
            for (var i = 1; i <= (int)range.Count; i++) selectedIds.Add((int)range.Item(i).Id);
            var others = new List<Box>();
            for (var i = 1; i <= (int)slide.Shapes.Count; i++)
            {
                dynamic shape = slide.Shapes.Item(i);
                if (!selectedIds.Contains((int)shape.Id)) others.Add(Read(shape, i));
            }
            for (var i = 1; i <= (int)range.Count; i++)
            {
                dynamic shape = range.Item(i);
                var box = Read(shape, i);
                Geometry.Snap(box, others, direction, slideWidth, slideHeight);
                shape.Left = box.Left;
                shape.Top = box.Top;
            }
        });

        public static void SetTextBoxMode(int mode) => Run((app, range) =>
        {
            for (var i = 1; i <= (int)range.Count; i++)
            {
                dynamic shape = range.Item(i);
                try { if ((int)shape.HasTextFrame != 0) shape.TextFrame2.AutoSize = mode; } catch { }
            }
        });

        public static void SetWordWrap(bool enabled) => Run((app, range) =>
        {
            for (var i = 1; i <= (int)range.Count; i++)
            {
                dynamic shape = range.Item(i);
                try { if ((int)shape.HasTextFrame != 0) shape.TextFrame2.WordWrap = enabled ? -1 : 0; } catch { }
            }
        });

        public static void SetMargins(decimal leftCm, decimal rightCm, decimal topCm, decimal bottomCm)
        {
            const float pointsPerCentimeter = 28.3464567f;
            Run((app, range) =>
            {
                for (var i = 1; i <= (int)range.Count; i++)
                {
                    dynamic shape = range.Item(i);
                    try
                    {
                        if ((int)shape.HasTextFrame == 0) continue;
                        shape.TextFrame.MarginLeft = (float)leftCm * pointsPerCentimeter;
                        shape.TextFrame.MarginRight = (float)rightCm * pointsPerCentimeter;
                        shape.TextFrame.MarginTop = (float)topCm * pointsPerCentimeter;
                        shape.TextFrame.MarginBottom = (float)bottomCm * pointsPerCentimeter;
                    }
                    catch { }
                }
            });
        }

        public static void SplitText(bool byParagraph) => Run((app, range) =>
        {
            dynamic source = range.Item(1);
            if ((int)source.HasTextFrame == 0) throw new InvalidOperationException("请选择包含文字的文本框。");
            string text = source.TextFrame.TextRange.Text;
            if (byParagraph)
            {
                var parts = text.Split(new[] { "\r\n", "\r" }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 2) throw new InvalidOperationException("文本中没有可拆分的段落。");
                float top = source.Top;
                foreach (var part in parts)
                {
                    dynamic duplicate = source.Duplicate().Item(1);
                    duplicate.TextFrame.TextRange.Text = part.TrimEnd();
                    duplicate.TextFrame.AutoSize = 1;
                    duplicate.Left = source.Left;
                    duplicate.Top = top;
                    duplicate.Width = source.Width;
                    top += duplicate.Height;
                }
            }
            else
            {
                var lines = text.Replace("\r\n", "\r").Replace("\n", "\r").Replace("\v", "\r").Split('\r');
                var characterCount = lines.Sum(line => new StringInfo(line).LengthInTextElements);
                if (characterCount < 2) throw new InvalidOperationException("文本中没有可拆分的单字。");
                float top = source.Top;
                foreach (var line in lines)
                {
                    float left = source.Left;
                    float lineHeight = 0f;
                    var characters = StringInfo.GetTextElementEnumerator(line);
                    while (characters.MoveNext())
                    {
                        dynamic duplicate = source.Duplicate().Item(1);
                        duplicate.TextFrame.TextRange.Text = characters.GetTextElement();
                        duplicate.TextFrame.AutoSize = 1;
                        duplicate.Left = left;
                        duplicate.Top = top;
                        left += duplicate.Width;
                        lineHeight = Math.Max(lineHeight, (float)duplicate.Height);
                    }
                    top += lineHeight;
                }
            }
            source.Delete();
        });

        public static void CombineText(bool withLineBreaks) => Run((app, range) =>
        {
            var items = new List<dynamic>();
            for (var i = 1; i <= (int)range.Count; i++) items.Add(range.Item(i));
            items = items.OrderBy(x => (float)x.Top).ThenBy(x => (float)x.Left).ToList();
            dynamic anchor = items[0];
            var texts = items.Select(x =>
            {
                try { return ((string)x.TextFrame.TextRange.Text).Trim(); }
                catch { return ""; }
            }).Where(x => x.Length > 0);
            anchor.TextFrame.TextRange.Text = string.Join(withLineBreaks ? "\r" : "", texts);
            for (var i = 1; i < items.Count; i++) items[i].Delete();
        }, 2);

        private static Box Read(dynamic shape, int id) => new Box
        {
            Id = id,
            Left = (float)shape.Left,
            Top = (float)shape.Top,
            Width = (float)shape.Width,
            Height = (float)shape.Height
        };
    }
}
