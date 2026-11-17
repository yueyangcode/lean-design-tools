using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace LeanDesignTools
{
    public enum AlignMode { Left, Center, Right, Top, Middle, Bottom }
    public enum StylePart { Width, Height, Position, Corner, Font, Fill, Line }

    public static class ShapeOperations
    {
        private static StyleSnapshot _capturedStyle;

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

        public static bool CaptureStyle()
        {
            try
            {
                dynamic app = DesignPane.PowerPoint;
                if (app == null) throw new InvalidOperationException("PowerPoint 尚未连接。");
                dynamic selection = app.ActiveWindow.Selection;
                if (((int)selection.Type != 2 && (int)selection.Type != 3) || (int)selection.ShapeRange.Count < 1)
                    throw new InvalidOperationException("请选择一个基准对象。");
                _capturedStyle = ReadStyle(selection.ShapeRange.Item(1));
                return true;
            }
            catch (Exception ex)
            {
                var message = ex is InvalidOperationException ? ex.Message : "吸取失败：" + ex.Message;
                MessageBox.Show(message, "设计工具", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
        }

        public static int SelectionCount()
        {
            try
            {
                dynamic selection = DesignPane.PowerPoint?.ActiveWindow?.Selection;
                return selection != null && ((int)selection.Type == 2 || (int)selection.Type == 3)
                    ? (int)selection.ShapeRange.Count : 0;
            }
            catch { return 0; }
        }

        public static string SingleSelectionKey()
        {
            try
            {
                dynamic app = DesignPane.PowerPoint;
                dynamic selection = app?.ActiveWindow?.Selection;
                if (selection == null || ((int)selection.Type != 2 && (int)selection.Type != 3) || (int)selection.ShapeRange.Count != 1) return null;
                dynamic shape = selection.ShapeRange.Item(1);
                return ((int)app.ActiveWindow.View.Slide.SlideID).ToString(CultureInfo.InvariantCulture) + ":" +
                    ((int)shape.Id).ToString(CultureInfo.InvariantCulture);
            }
            catch { return null; }
        }

        public static void ApplyStyle(StylePart part) => Run((app, range) =>
        {
            var multiple = (int)range.Count >= 2;
            var style = multiple ? ReadStyle(range.Item(1)) : _capturedStyle;
            if (style == null) throw new InvalidOperationException("请先吸取基准样式。");

            var applied = 0;
            for (var i = multiple ? 2 : 1; i <= (int)range.Count; i++)
                if (TryApplyStyle(range.Item(i), style, part)) applied++;
            if (applied == 0) throw new InvalidOperationException("所选对象不支持此格式。");
        });

        public static bool ApplyCapturedStyle(StylePart part)
        {
            var applied = false;
            Run((app, range) =>
            {
                if ((int)range.Count != 1) throw new InvalidOperationException("请选择一个目标对象。");
                if (_capturedStyle == null) throw new InvalidOperationException("请先吸取基准样式。");
                applied = TryApplyStyle(range.Item(1), _capturedStyle, part);
                if (!applied) throw new InvalidOperationException("所选对象不支持此格式。");
            });
            return applied;
        }

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

        private static StyleSnapshot ReadStyle(dynamic shape)
        {
            var style = new StyleSnapshot
            {
                Left = (float)shape.Left,
                Top = (float)shape.Top,
                Width = (float)shape.Width,
                Height = (float)shape.Height
            };

            try
            {
                if (SupportsCorner(shape) && (int)shape.Adjustments.Count > 0)
                    style.CornerRadius = (float)shape.Adjustments[1] * Math.Min(style.Width, style.Height);
            }
            catch { }

            try
            {
                if ((int)shape.HasTextFrame != 0)
                {
                    dynamic textRange = shape.TextFrame.TextRange;
                    dynamic font;
                    try { font = textRange.Characters(1, 1).Font; }
                    catch { font = textRange.Font; }
                    style.Font = new FontSnapshot
                    {
                        Name = (string)font.Name,
                        NameFarEast = (string)font.NameFarEast,
                        Size = (float)font.Size,
                        Bold = (int)font.Bold,
                        Italic = (int)font.Italic,
                        Color = ReadColor(font.Color)
                    };
                }
            }
            catch { }

            try
            {
                var visible = (int)shape.Fill.Visible != 0;
                style.Fill = new FillSnapshot { Visible = visible };
                if (visible && (int)shape.Fill.Type == 1)
                {
                    style.Fill.Supported = true;
                    style.Fill.Color = ReadColor(shape.Fill.ForeColor);
                    style.Fill.Transparency = (float)shape.Fill.Transparency;
                }
                else if (!visible) style.Fill.Supported = true;
            }
            catch { }

            try
            {
                var visible = (int)shape.Line.Visible != 0;
                style.Line = new LineSnapshot { Visible = visible, Supported = true };
                if (visible)
                {
                    style.Line.Color = ReadColor(shape.Line.ForeColor);
                    style.Line.Weight = (float)shape.Line.Weight;
                    style.Line.DashStyle = (int)shape.Line.DashStyle;
                    style.Line.Transparency = (float)shape.Line.Transparency;
                }
            }
            catch { }

            return style;
        }

        private static bool TryApplyStyle(dynamic shape, StyleSnapshot style, StylePart part)
        {
            try
            {
                switch (part)
                {
                    case StylePart.Width:
                        SetIndependentDimension(shape, style.Width, true);
                        return true;
                    case StylePart.Height:
                        SetIndependentDimension(shape, style.Height, false);
                        return true;
                    case StylePart.Position:
                        shape.Left = style.Left;
                        shape.Top = style.Top;
                        return true;
                    case StylePart.Corner:
                        if (!style.CornerRadius.HasValue || !SupportsCorner(shape) || (int)shape.Adjustments.Count == 0) return false;
                        var minimum = Math.Min((float)shape.Width, (float)shape.Height);
                        if (minimum <= 0) return false;
                        shape.Adjustments[1] = Math.Max(0f, Math.Min(0.5f, style.CornerRadius.Value / minimum));
                        return true;
                    case StylePart.Font:
                        if (style.Font == null || (int)shape.HasTextFrame == 0) return false;
                        ApplyFont(shape.TextFrame.TextRange.Font, style.Font);
                        return true;
                    case StylePart.Fill:
                        if (style.Fill == null || !style.Fill.Supported) return false;
                        ApplyFill(shape.Fill, style.Fill);
                        return true;
                    case StylePart.Line:
                        if (style.Line == null || !style.Line.Supported) return false;
                        ApplyLine(shape.Line, style.Line);
                        return true;
                    default:
                        return false;
                }
            }
            catch { return false; }
        }

        private static void SetIndependentDimension(dynamic shape, float value, bool width)
        {
            var locked = (int)shape.LockAspectRatio;
            shape.LockAspectRatio = 0;
            if (width) shape.Width = value; else shape.Height = value;
            shape.LockAspectRatio = locked;
        }

        private static bool SupportsCorner(dynamic shape)
        {
            try
            {
                var type = (int)shape.AutoShapeType;
                return type == 5 || (type >= 151 && type <= 154);
            }
            catch { return false; }
        }

        private static ColorSnapshot ReadColor(dynamic color)
        {
            var snapshot = new ColorSnapshot { Rgb = (int)color.RGB };
            try
            {
                var theme = (int)color.ObjectThemeColor;
                if (theme > 0)
                {
                    snapshot.HasTheme = true;
                    snapshot.Theme = theme;
                    snapshot.TintAndShade = (float)color.TintAndShade;
                }
            }
            catch { }
            return snapshot;
        }

        private static void ApplyColor(dynamic target, ColorSnapshot color)
        {
            if (color.HasTheme)
            {
                try
                {
                    target.ObjectThemeColor = color.Theme;
                    target.TintAndShade = color.TintAndShade;
                    return;
                }
                catch { }
            }
            target.RGB = color.Rgb;
        }

        private static void ApplyFont(dynamic target, FontSnapshot font)
        {
            target.Name = font.Name;
            target.NameFarEast = font.NameFarEast;
            target.Size = font.Size;
            target.Bold = font.Bold;
            target.Italic = font.Italic;
            ApplyColor(target.Color, font.Color);
        }

        private static void ApplyFill(dynamic target, FillSnapshot fill)
        {
            target.Visible = fill.Visible ? -1 : 0;
            if (!fill.Visible) return;
            target.Solid();
            ApplyColor(target.ForeColor, fill.Color);
            target.Transparency = fill.Transparency;
        }

        private static void ApplyLine(dynamic target, LineSnapshot line)
        {
            target.Visible = line.Visible ? -1 : 0;
            if (!line.Visible) return;
            ApplyColor(target.ForeColor, line.Color);
            target.Weight = line.Weight;
            target.DashStyle = line.DashStyle;
            target.Transparency = line.Transparency;
        }

        private sealed class StyleSnapshot
        {
            public float Left;
            public float Top;
            public float Width;
            public float Height;
            public float? CornerRadius;
            public FontSnapshot Font;
            public FillSnapshot Fill;
            public LineSnapshot Line;
        }

        private sealed class FontSnapshot
        {
            public string Name;
            public string NameFarEast;
            public float Size;
            public int Bold;
            public int Italic;
            public ColorSnapshot Color;
        }

        private sealed class FillSnapshot
        {
            public bool Supported;
            public bool Visible;
            public ColorSnapshot Color;
            public float Transparency;
        }

        private sealed class LineSnapshot
        {
            public bool Supported;
            public bool Visible;
            public ColorSnapshot Color;
            public float Weight;
            public int DashStyle;
            public float Transparency;
        }

        private sealed class ColorSnapshot
        {
            public int Rgb;
            public bool HasTheme;
            public int Theme;
            public float TintAndShade;
        }

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
