using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace LeanDesignTools
{
    [ComVisible(true)]
    [Guid("AA7C4797-FD89-49BC-8E1A-A0C63A3788F0")]
    [ProgId("LeanDesignTools.Pane")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class DesignPane : UserControl
    {
        internal static dynamic PowerPoint;
        private SmoothFlowLayoutPanel _content;
        private readonly List<SmoothFlowLayoutPanel> _contents = new List<SmoothFlowLayoutPanel>();
        private readonly List<Panel> _pages = new List<Panel>();
        private readonly List<NavButton> _navButtons = new List<NavButton>();
        private readonly TableLayoutPanel _nav;
        private readonly Panel _pageHost;
        private readonly ToolTip _tips = new ToolTip { ShowAlways = true, InitialDelay = 250, ReshowDelay = 100, AutoPopDelay = 8000 };
        private readonly List<ToolButton> _tools = new List<ToolButton>();
        private readonly List<Control> _shapeControls = new List<Control>();
        private readonly Timer _stateTimer;
        private CentimeterBox _marginLeft;
        private CentimeterBox _marginRight;
        private CentimeterBox _marginTop;
        private CentimeterBox _marginBottom;
        private CentimeterBox _objectWidth;
        private CentimeterBox _objectHeight;
        private AccentRadioButton _freeMode;
        private AccentRadioButton _fitShapeMode;
        private AccentRadioButton _shrinkMode;
        private AccentCheckBox _wordWrap;

        public DesignPane()
        {
            BackColor = Color.White;
            Font = new Font("Microsoft YaHei UI", 9f);
            TabStop = true;
            _nav = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 33,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = Color.White,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            for (var i = 0; i < 3; i++) _nav.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / 3));
            _pageHost = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
            Controls.Add(_pageHost);
            Controls.Add(_nav);
            BuildUi();
            SelectPage(0);
            Load += (s, e) =>
            {
                var resetScroll = new Timer { Interval = 50 };
                resetScroll.Tick += (sender, args) =>
                {
                    resetScroll.Stop();
                    Select();
                    foreach (var content in _contents) content.AutoScrollPosition = Point.Empty;
                    resetScroll.Dispose();
                };
                resetScroll.Start();
            };
            _stateTimer = new Timer { Interval = 300 };
            _stateTimer.Tick += (s, e) => UpdateButtonStates();
            UpdateButtonStates();
            _stateTimer.Start();
        }

        internal static void ShowError(string title, Exception ex) =>
            MessageBox.Show(ex.Message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);

        private void BuildUi()
        {
            BeginPage("布局");
            AddSection("对齐",
                Tool("左对齐", IconKind.AlignLeft, () => ShapeOperations.Align(AlignMode.Left), 2),
                Tool("水平居中", IconKind.AlignCenter, () => ShapeOperations.Align(AlignMode.Center), 2),
                Tool("右对齐", IconKind.AlignRight, () => ShapeOperations.Align(AlignMode.Right), 2),
                Tool("顶对齐", IconKind.AlignTop, () => ShapeOperations.Align(AlignMode.Top), 2),
                Tool("垂直居中", IconKind.AlignMiddle, () => ShapeOperations.Align(AlignMode.Middle), 2),
                Tool("底对齐", IconKind.AlignBottom, () => ShapeOperations.Align(AlignMode.Bottom), 2),
                Tool("水平分布", IconKind.DistributeH, () => ShapeOperations.Distribute(true), 3),
                Tool("垂直分布", IconKind.DistributeV, () => ShapeOperations.Distribute(false), 3));

            AddSection("大小",
                Tool("等宽", IconKind.SameWidth, () => ShapeOperations.MatchSize(true, false), 2),
                Tool("等高", IconKind.SameHeight, () => ShapeOperations.MatchSize(false, true), 2),
                Tool("等大", IconKind.SameSize, () => ShapeOperations.MatchSize(true, true), 2),
                Tool("横向铺满", IconKind.StretchH, () => ShapeOperations.Stretch(true)),
                Tool("纵向铺满", IconKind.StretchV, () => ShapeOperations.Stretch(false)));

            AddSection("参考线布局",
                Tool("贴左", IconKind.DockLeft, () => ShapeOperations.Dock(-1, 2)),
                Tool("水平居中", IconKind.DockCenterH, () => ShapeOperations.Dock(0, 2)),
                Tool("贴右", IconKind.DockRight, () => ShapeOperations.Dock(1, 2)),
                Tool("贴顶", IconKind.DockTop, () => ShapeOperations.Dock(2, -1)),
                Tool("垂直居中", IconKind.DockCenterV, () => ShapeOperations.Dock(2, 0)),
                Tool("贴底", IconKind.DockBottom, () => ShapeOperations.Dock(2, 1)),
                Tool("幻灯片居中", IconKind.SlideCenter, () => ShapeOperations.Dock(0, 0)));

            BeginPage("形状");
            AddSection("选择",
                Tool("全选形状", IconKind.SelectAll, ShapeOperations.SelectAll, 0),
                Tool("组合", IconKind.Group, ShapeOperations.Group, 2),
                Tool("取消组合", IconKind.Ungroup, ShapeOperations.Ungroup),
                Tool("置于顶层", IconKind.BringFront, () => ShapeOperations.ZOrder(0)),
                Tool("上移一层", IconKind.Forward, () => ShapeOperations.ZOrder(2)),
                Tool("下移一层", IconKind.Backward, () => ShapeOperations.ZOrder(3)),
                Tool("置于底层", IconKind.SendBack, () => ShapeOperations.ZOrder(1)),
                Tool("选择窗格", IconKind.Eye, () => ShapeOperations.ExecuteMso("SelectionPane", 0), 0));

            AddSection("矢量",
                Tool("联合", IconKind.Union, () => ShapeOperations.ExecuteMso("ShapesUnion"), 2),
                Tool("组合", IconKind.Combine, () => ShapeOperations.ExecuteMso("ShapesCombine"), 2),
                Tool("拆分", IconKind.Fragment, () => ShapeOperations.ExecuteMso("ShapesFragment"), 2),
                Tool("相交", IconKind.Intersect, () => ShapeOperations.ExecuteMso("ShapesIntersect"), 2),
                Tool("剪除", IconKind.Subtract, () => ShapeOperations.ExecuteMso("ShapesSubtract"), 2));

            AddSection("剪贴板", 2,
                Tool("复制位置（可跨幻灯片）", IconKind.CopyPosition, () => ShapeOperations.CopyGeometry(true, false, false)),
                Tool("粘贴位置（可跨幻灯片）", IconKind.PastePosition, () => ShapeOperations.PasteGeometry(true, false, false)),
                Tool("复制宽度（可跨幻灯片）", IconKind.CopyWidth, () => ShapeOperations.CopyGeometry(false, true, false)),
                Tool("粘贴宽度（可跨幻灯片）", IconKind.PasteWidth, () => ShapeOperations.PasteGeometry(false, true, false)),
                Tool("复制高度（可跨幻灯片）", IconKind.CopyHeight, () => ShapeOperations.CopyGeometry(false, false, true)),
                Tool("粘贴高度（可跨幻灯片）", IconKind.PasteHeight, () => ShapeOperations.PasteGeometry(false, false, true)),
                Tool("复制大小（可跨幻灯片）", IconKind.CopySize, () => ShapeOperations.CopyGeometry(false, true, true)),
                Tool("粘贴大小（可跨幻灯片）", IconKind.PasteSize, () => ShapeOperations.PasteGeometry(false, true, true)));

            AddSection("吸附", 2,
                Tool("向左吸附", IconKind.SnapLeft, () => ShapeOperations.Snap(SnapDirection.Left)),
                Tool("向右吸附", IconKind.SnapRight, () => ShapeOperations.Snap(SnapDirection.Right)),
                Tool("向上吸附", IconKind.SnapUp, () => ShapeOperations.Snap(SnapDirection.Up)),
                Tool("向下吸附", IconKind.SnapDown, () => ShapeOperations.Snap(SnapDirection.Down)));

            AddSection("旋转", 2,
                Tool("向左旋转 90°", IconKind.RotateLeft, () => ShapeOperations.Rotate(-90)),
                Tool("向右旋转 90°", IconKind.RotateRight, () => ShapeOperations.Rotate(90)),
                Tool("水平翻转", IconKind.FlipH, () => ShapeOperations.Flip(true)),
                Tool("垂直翻转", IconKind.FlipV, () => ShapeOperations.Flip(false)));

            BeginPage("对象");
            AddSection("拆合", 2,
                Tool("单字拆分", IconKind.SplitParagraph, () => ShapeOperations.SplitText(false)),
                Tool("单字合并", IconKind.CombineParagraph, () => ShapeOperations.CombineText(false), 2),
                Tool("按段拆分", IconKind.SplitLine, () => ShapeOperations.SplitText(true)),
                Tool("段落合并", IconKind.CombineLine, () => ShapeOperations.CombineText(true), 2));

            AddObjectSizeSection();
            AddTextBoxSection();
            AddMarginSection();
        }

        private void BeginPage(string title)
        {
            var page = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
            var content = new SmoothFlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                BackColor = Color.White,
                Padding = new Padding(10, 4, 4, 10)
            };
            content.Resize += (s, e) => ResizeSections(content);
            page.Controls.Add(content);
            _pageHost.Controls.Add(page);

            var index = _pages.Count;
            var button = new NavButton { Text = title, Dock = DockStyle.Fill, Margin = new Padding(0), AccessibleName = title + "页签" };
            button.Click += (s, e) => SelectPage(index);
            _nav.Controls.Add(button, index, 0);
            _pages.Add(page);
            _contents.Add(content);
            _navButtons.Add(button);
            _content = content;
        }

        private void SelectPage(int index)
        {
            for (var i = 0; i < _pages.Count; i++)
            {
                _pages[i].Visible = i == index;
                _navButtons[i].Selected = i == index;
            }
            _pages[index].BringToFront();
            _contents[index].AutoScrollPosition = Point.Empty;
        }

        private ToolButton Tool(string tip, IconKind icon, Action action, int minimumSelection = 1)
        {
            var button = new ToolButton
            {
                Icon = icon,
                MinimumSelection = minimumSelection,
                AccessibleName = tip,
                TabStop = minimumSelection > 0
            };
            button.Click += (s, e) => action();
            _tips.SetToolTip(button, tip);
            _tools.Add(button);
            _shapeControls.Add(button);
            return button;
        }

        private void AddSection(string title, params ToolButton[] buttons)
        {
            AddSection(title, 3, buttons);
        }

        private void AddSection(string title, int columns, params ToolButton[] buttons)
        {
            var rows = (buttons.Length + columns - 1) / columns;
            Panel header;
            Panel body;
            CreateSection(title, rows * 29 + 8, title + "工具", out header, out body);
            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = columns,
                RowCount = rows,
                Padding = new Padding(2, 2, 2, 4),
                BackColor = Color.White
            };
            for (var i = 0; i < columns; i++) grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / columns));
            foreach (var button in buttons) grid.Controls.Add(button);
            body.Controls.Add(grid);
        }

        private void CreateSection(string title, int bodyHeight, string tip, out Panel header, out Panel body)
        {
            var panel = new Panel
            {
                Height = 30 + bodyHeight,
                Margin = new Padding(0),
                BackColor = Color.White
            };
            body = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
            header = new Panel { Dock = DockStyle.Top, Height = 30, BackColor = Color.White };
            var separator = new Panel { Dock = DockStyle.Bottom, Height = 1, BackColor = Color.FromArgb(238, 238, 238) };
            var label = new Label
            {
                Text = title + "  ⓘ",
                ForeColor = Color.FromArgb(100, 100, 100),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(4, 0, 0, 0),
                AccessibleName = title
            };
            _tips.SetToolTip(label, tip);
            header.Controls.Add(label);
            panel.Controls.Add(body);
            panel.Controls.Add(header);
            panel.Controls.Add(separator);
            separator.BringToFront();
            _content.Controls.Add(panel);
        }

        private void AddObjectSizeSection()
        {
            Panel header;
            Panel body;
            CreateSection("对象尺寸", 38, "调整所选图片或形状的宽度和高度（厘米）", out header, out body);
            var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Padding = new Padding(1, 5, 1, 5) };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            _objectWidth = DimensionBox("宽", true);
            _objectHeight = DimensionBox("高", false);
            grid.Controls.Add(WrapDimensionBox(_objectWidth, true));
            grid.Controls.Add(WrapDimensionBox(_objectHeight, false));
            body.Controls.Add(grid);
        }

        private CentimeterBox DimensionBox(string name, bool width)
        {
            var box = new CentimeterBox { DecimalPlaces = 2, Increment = 0.1m, Minimum = 0.01m, Maximum = 1000m, Value = 1m, Tag = name, AccessibleName = name + "度（厘米）" };
            box.ValueCommitted += (s, e) => ShapeOperations.SetDimensions(width ? (decimal?)box.Value : null, width ? null : (decimal?)box.Value);
            _tips.SetToolTip(box, name + "度（厘米）");
            _tips.SetToolTip(box.Editor, name + "度（厘米）");
            return box;
        }

        private Control WrapDimensionBox(CentimeterBox box, bool width)
        {
            var panel = new Panel { Height = 28, Margin = new Padding(0), Padding = new Padding(16, 1, 0, 4) };
            var glyph = new DimensionGlyph { Horizontal = width, Location = new Point(0, 3), Size = new Size(15, 17) };
            box.Dock = DockStyle.Fill;
            box.AvailabilityChanged += (s, e) => glyph.Available = box.Available;
            _tips.SetToolTip(glyph, width ? "宽度（厘米）" : "高度（厘米）");
            panel.Controls.Add(box);
            panel.Controls.Add(glyph);
            glyph.BringToFront();
            return panel;
        }

        private void AddTextBoxSection()
        {
            Panel header;
            Panel body;
            CreateSection("文本框", 62, "自由调整：可以自由调整文本框大小\r\n根据文字：自动调整文本框大小适应文字\r\n溢出缩排：自动调整文字大小适应文本框", out header, out body);
            var clear = Tool("清空文本", IconKind.ClearText, ShapeOperations.ClearText);
            clear.Size = new Size(28, 28); clear.Dock = DockStyle.Right; header.Controls.Add(clear); clear.BringToFront();
            var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, Padding = new Padding(2, 2, 2, 4) };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            _freeMode = ModeRadio("自由调整", "可以自由调整文本框大小", () => ShapeOperations.SetTextBoxMode(0));
            _fitShapeMode = ModeRadio("根据文字", "自动调整文本框大小适应文字", () => ShapeOperations.SetTextBoxMode(1));
            _shrinkMode = ModeRadio("溢出缩排", "自动调整文字大小适应文本框", () => ShapeOperations.SetTextBoxMode(2));
            _wordWrap = new AccentCheckBox { Text = "自动换行", Dock = DockStyle.Fill, ForeColor = Color.FromArgb(100, 100, 100), Margin = new Padding(2, 1, 0, 0) };
            _wordWrap.CheckedChanged += (s, e) => { if (_wordWrap.Focused) ShapeOperations.SetWordWrap(_wordWrap.Checked); };
            _tips.SetToolTip(_wordWrap, "切换文本框自动换行");
            _shapeControls.Add(_wordWrap);
            grid.Controls.Add(_freeMode); grid.Controls.Add(_fitShapeMode); grid.Controls.Add(_shrinkMode); grid.Controls.Add(_wordWrap);
            body.Controls.Add(grid);
        }

        private AccentRadioButton ModeRadio(string text, string tip, Action action)
        {
            var radio = new AccentRadioButton { Text = text, AutoCheck = false, Dock = DockStyle.Fill, ForeColor = Color.FromArgb(100, 100, 100), Margin = new Padding(2, 1, 0, 0) };
            radio.Click += (s, e) => action();
            _tips.SetToolTip(radio, tip);
            _shapeControls.Add(radio);
            return radio;
        }

        private void AddMarginSection()
        {
            Panel header;
            Panel body;
            CreateSection("文本框边距", 82, "设置文字与文本框边缘的距离（厘米）", out header, out body);
            var clear = Tool("删除边距", IconKind.ClearMargins, () => SetMarginValues(0m, 0m, 0m, 0m));
            clear.Size = new Size(28, 28);
            var reset = Tool("重置边距", IconKind.ResetMargins, () => SetMarginValues(0.25m, 0.25m, 0.13m, 0.13m));
            reset.Size = new Size(28, 28);
            var actions = new FlowLayoutPanel { Dock = DockStyle.Right, Width = 56, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Margin = new Padding(0), Padding = new Padding(0) };
            actions.Controls.Add(clear); actions.Controls.Add(reset); header.Controls.Add(actions); actions.BringToFront();

            var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, Padding = new Padding(1, 4, 1, 4) };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            _marginLeft = MarginBox("左", 0.25m); _marginRight = MarginBox("右", 0.25m); _marginTop = MarginBox("上", 0.13m); _marginBottom = MarginBox("下", 0.13m);
            grid.Controls.Add(WrapMarginBox(_marginLeft, MarginDirection.Left)); grid.Controls.Add(WrapMarginBox(_marginRight, MarginDirection.Right));
            grid.Controls.Add(WrapMarginBox(_marginTop, MarginDirection.Top)); grid.Controls.Add(WrapMarginBox(_marginBottom, MarginDirection.Bottom));
            body.Controls.Add(grid);
        }

        private CentimeterBox MarginBox(string name, decimal value)
        {
            var box = new CentimeterBox { DecimalPlaces = 2, Increment = 0.05m, Minimum = 0, Maximum = 10, Value = value, Width = 74, Tag = name, AccessibleName = name + "边距（厘米）" };
            box.ValueCommitted += (s, e) => ApplyMargins();
            _shapeControls.Add(box);
            _tips.SetToolTip(box, name + "边距（厘米）");
            _tips.SetToolTip(box.Editor, name + "边距（厘米）");
            return box;
        }

        private static Control WrapMarginBox(CentimeterBox box, MarginDirection direction)
        {
            var panel = new Panel { Height = 29, Margin = new Padding(0), Padding = new Padding(16, 2, 0, 4) };
            var glyph = new MarginDirectionGlyph { Direction = direction, Location = new Point(0, 4), Size = new Size(15, 17) };
            box.Dock = DockStyle.Fill;
            box.AvailabilityChanged += (s, e) => glyph.Available = box.Available;
            panel.Controls.Add(box);
            panel.Controls.Add(glyph);
            glyph.BringToFront();
            return panel;
        }

        private void SetMarginValues(decimal left, decimal right, decimal top, decimal bottom)
        {
            _marginLeft.Value = left; _marginRight.Value = right; _marginTop.Value = top; _marginBottom.Value = bottom; ApplyMargins();
        }

        private void ApplyMargins()
        {
            if (PowerPoint == null || !_marginLeft.Available) return;
            ShapeOperations.SetMargins(_marginLeft.Value, _marginRight.Value, _marginTop.Value, _marginBottom.Value);
        }

        private static void ResizeSections(SmoothFlowLayoutPanel content)
        {
            var width = Math.Max(150, content.ClientSize.Width - content.Padding.Horizontal - 2);
            foreach (Control control in content.Controls) control.Width = width;
        }

        private void UpdateButtonStates()
        {
            var count = 0;
            dynamic firstShape = null;
            try
            {
                dynamic selection = PowerPoint?.ActiveWindow?.Selection;
                if (selection != null && ((int)selection.Type == 2 || (int)selection.Type == 3))
                {
                    dynamic range = selection.ShapeRange;
                    count = (int)range.Count;
                    if (count > 0) firstShape = range.Item(1);
                }
            }
            catch { }
            foreach (var tool in _tools) tool.Available = count >= tool.MinimumSelection;
            var hasTextFrame = false;
            try { hasTextFrame = firstShape != null && (int)firstShape.HasTextFrame != 0; }
            catch { }
            _freeMode.Available = hasTextFrame;
            _fitShapeMode.Available = hasTextFrame;
            _shrinkMode.Available = hasTextFrame;
            _wordWrap.Available = hasTextFrame;
            _objectWidth.Available = firstShape != null;
            _objectHeight.Available = firstShape != null;
            _marginLeft.Available = hasTextFrame;
            _marginRight.Available = hasTextFrame;
            _marginTop.Available = hasTextFrame;
            _marginBottom.Available = hasTextFrame;
            if (firstShape != null && !_objectWidth.ContainsFocus && !_objectHeight.ContainsFocus)
            {
                const decimal centimetersPerPoint = 0.035277778m;
                SetDimensionDisplay(_objectWidth, (decimal)(float)firstShape.Width * centimetersPerPoint);
                SetDimensionDisplay(_objectHeight, (decimal)(float)firstShape.Height * centimetersPerPoint);
            }
            if (hasTextFrame) SyncTextBoxState(firstShape);
        }

        private void SyncTextBoxState(dynamic shape)
        {
            try
            {
                if ((int)shape.HasTextFrame == 0) return;
                var autoSize = (int)shape.TextFrame2.AutoSize;
                _freeMode.Checked = autoSize == 0;
                _fitShapeMode.Checked = autoSize == 1;
                _shrinkMode.Checked = autoSize == 2;
                _wordWrap.Checked = (int)shape.TextFrame2.WordWrap != 0;

                if (_marginLeft.ContainsFocus || _marginRight.ContainsFocus || _marginTop.ContainsFocus || _marginBottom.ContainsFocus) return;
                const decimal centimetersPerPoint = 0.035277778m;
                SetMarginDisplay(_marginLeft, (decimal)(float)shape.TextFrame.MarginLeft * centimetersPerPoint);
                SetMarginDisplay(_marginRight, (decimal)(float)shape.TextFrame.MarginRight * centimetersPerPoint);
                SetMarginDisplay(_marginTop, (decimal)(float)shape.TextFrame.MarginTop * centimetersPerPoint);
                SetMarginDisplay(_marginBottom, (decimal)(float)shape.TextFrame.MarginBottom * centimetersPerPoint);
            }
            catch { }
        }

        private static void SetMarginDisplay(CentimeterBox box, decimal value)
        {
            box.Value = Math.Max(box.Minimum, Math.Min(box.Maximum, Math.Round(value, box.DecimalPlaces)));
        }

        private static void SetDimensionDisplay(CentimeterBox box, decimal value)
        {
            box.Value = Math.Max(box.Minimum, Math.Min(box.Maximum, Math.Round(value, box.DecimalPlaces)));
        }
    }

    internal enum MarginDirection { Left, Right, Top, Bottom }

    internal sealed class DimensionGlyph : Control
    {
        public bool Horizontal { get; set; }
        private bool _available = true;
        public bool Available { get => _available; set { _available = value; Invalidate(); } }

        public DimensionGlyph()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            using (var pen = new Pen(Available ? Color.FromArgb(244, 78, 72) : Color.FromArgb(205, 205, 205), 1f))
            {
                pen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                if (Horizontal)
                {
                    e.Graphics.DrawLine(pen, 2, 4, 2, 13); e.Graphics.DrawLine(pen, 12, 4, 12, 13);
                    e.Graphics.DrawLine(pen, 3, 8, 11, 8);
                    e.Graphics.DrawLine(pen, 3, 8, 6, 5); e.Graphics.DrawLine(pen, 3, 8, 6, 11);
                    e.Graphics.DrawLine(pen, 11, 8, 8, 5); e.Graphics.DrawLine(pen, 11, 8, 8, 11);
                }
                else
                {
                    e.Graphics.DrawLine(pen, 3, 2, 11, 2); e.Graphics.DrawLine(pen, 3, 14, 11, 14);
                    e.Graphics.DrawLine(pen, 7, 3, 7, 13);
                    e.Graphics.DrawLine(pen, 7, 3, 4, 6); e.Graphics.DrawLine(pen, 7, 3, 10, 6);
                    e.Graphics.DrawLine(pen, 7, 13, 4, 10); e.Graphics.DrawLine(pen, 7, 13, 10, 10);
                }
            }
        }
    }

    internal sealed class MarginDirectionGlyph : Control
    {
        public MarginDirection Direction { get; set; }
        private bool _available = true;
        public bool Available { get => _available; set { _available = value; Invalidate(); } }

        public MarginDirectionGlyph()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            using (var pen = new Pen(Available ? Color.FromArgb(244, 78, 72) : Color.FromArgb(205, 205, 205), 1f))
            {
                pen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                switch (Direction)
                {
                    case MarginDirection.Left:
                        e.Graphics.DrawLine(pen, 2, 4, 2, 13); e.Graphics.DrawLine(pen, 3, 8, 13, 8);
                        e.Graphics.DrawLine(pen, 3, 8, 6, 5); e.Graphics.DrawLine(pen, 3, 8, 6, 11); break;
                    case MarginDirection.Right:
                        e.Graphics.DrawLine(pen, 13, 4, 13, 13); e.Graphics.DrawLine(pen, 2, 8, 12, 8);
                        e.Graphics.DrawLine(pen, 12, 8, 9, 5); e.Graphics.DrawLine(pen, 12, 8, 9, 11); break;
                    case MarginDirection.Top:
                        e.Graphics.DrawLine(pen, 3, 2, 12, 2); e.Graphics.DrawLine(pen, 7, 3, 7, 14);
                        e.Graphics.DrawLine(pen, 7, 3, 4, 6); e.Graphics.DrawLine(pen, 7, 3, 10, 6); break;
                    case MarginDirection.Bottom:
                        e.Graphics.DrawLine(pen, 3, 14, 12, 14); e.Graphics.DrawLine(pen, 7, 2, 7, 13);
                        e.Graphics.DrawLine(pen, 7, 13, 4, 10); e.Graphics.DrawLine(pen, 7, 13, 10, 10); break;
                }
            }
        }
    }

    internal sealed class CentimeterBox : UserControl
    {
        private bool _available = true;
        private decimal _value;
        private readonly TextBox _editor;
        public event EventHandler AvailabilityChanged;
        public event EventHandler ValueCommitted;
        public int DecimalPlaces { get; set; } = 2;
        public decimal Increment { get; set; } = 0.05m;
        public decimal Minimum { get; set; }
        public decimal Maximum { get; set; } = 10m;
        public Control Editor => _editor;

        public CentimeterBox()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
            BackColor = Color.White;
            ForeColor = Color.FromArgb(100, 100, 100);
            Font = new Font("Microsoft YaHei UI", 8.25f);
            Height = 23;
            TabStop = true;
            _editor = new TextBox { BorderStyle = BorderStyle.None, BackColor = Color.White, ForeColor = ForeColor, Font = Font, TextAlign = HorizontalAlignment.Left };
            _editor.KeyDown += (s, e) =>
            {
                if (e.KeyCode != Keys.Enter) return;
                CommitEditor();
                e.SuppressKeyPress = true;
            };
            _editor.Leave += (s, e) => CommitEditor();
            Controls.Add(_editor);
            UpdateEditorBounds();
            UpdateText();
        }

        public decimal Value
        {
            get => _value;
            set { _value = Math.Max(Minimum, Math.Min(Maximum, value)); UpdateText(); }
        }

        public bool Available
        {
            get => _available;
            set
            {
                if (_available == value) return;
                _available = value;
                TabStop = value;
                ForeColor = value ? Color.FromArgb(100, 100, 100) : Color.FromArgb(190, 190, 190);
                _editor.ReadOnly = !value;
                _editor.TabStop = value;
                _editor.ForeColor = ForeColor;
                AvailabilityChanged?.Invoke(this, EventArgs.Empty);
                Invalidate();
            }
        }

        protected override void OnResize(EventArgs e) { base.OnResize(e); UpdateEditorBounds(); }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (!Available) return;
            if (e.X >= Width - 15)
            {
                Value += e.Y < Height / 2 ? Increment : -Increment;
                ValueCommitted?.Invoke(this, EventArgs.Empty);
            }
            else _editor.Focus();
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            if (!Available) return;
            Value += e.Delta > 0 ? Increment : -Increment;
            ValueCommitted?.Invoke(this, EventArgs.Empty);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var border = Available ? Color.FromArgb(232, 232, 232) : Color.FromArgb(242, 242, 242);
            ControlPaint.DrawBorder(e.Graphics, ClientRectangle, border, ButtonBorderStyle.Solid);
            var arrow = Available ? Color.FromArgb(90, 90, 90) : Color.FromArgb(205, 205, 205);
            using (var brush = new SolidBrush(arrow))
            {
                var x = Width - 7;
                e.Graphics.FillPolygon(brush, new[] { new Point(x, 6), new Point(x - 2, 8), new Point(x + 2, 8) });
                e.Graphics.FillPolygon(brush, new[] { new Point(x, 17), new Point(x - 2, 15), new Point(x + 2, 15) });
            }
        }

        private void UpdateEditorBounds()
        {
            if (_editor == null) return;
            _editor.SetBounds(4, 3, Math.Max(10, Width - 21), 17);
        }

        private void CommitEditor()
        {
            if (!Available) return;
            decimal parsed;
            var text = _editor.Text.Replace("cm", string.Empty).Trim();
            if (decimal.TryParse(text, out parsed)) Value = parsed;
            else UpdateText();
            ValueCommitted?.Invoke(this, EventArgs.Empty);
        }

        private void UpdateText()
        {
            if (_editor != null) _editor.Text = Value.ToString("F" + DecimalPlaces) + " cm";
        }
    }

    internal sealed class SmoothFlowLayoutPanel : FlowLayoutPanel
    {
        public SmoothFlowLayoutPanel()
        {
            DoubleBuffered = true;
            ResizeRedraw = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        }
    }

    internal sealed class NavButton : Button
    {
        private bool _selected;
        public bool Selected { get => _selected; set { _selected = value; Invalidate(); } }

        public NavButton()
        {
            BackColor = Color.White;
            Cursor = Cursors.Hand;
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            TabStop = true;
            MouseEnter += (s, e) => { BackColor = Color.FromArgb(247, 247, 247); Invalidate(); };
            MouseLeave += (s, e) => { BackColor = Color.White; Invalidate(); };
        }

        protected override bool ShowFocusCues => false;
        protected override bool ShowKeyboardCues => false;

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.Clear(BackColor);
            var color = Selected ? Color.FromArgb(244, 78, 72) : Color.FromArgb(100, 100, 100);
            TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, color, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
            if (Selected)
                using (var brush = new SolidBrush(Color.FromArgb(244, 78, 72)))
                    e.Graphics.FillRectangle(brush, 12, Height - 2, Math.Max(0, Width - 24), 2);
        }
    }

    internal sealed class AccentRadioButton : RadioButton
    {
        private bool _available = true;
        public bool Available { get => _available; set { if (_available == value) return; _available = value; TabStop = value; Invalidate(); } }

        public AccentRadioButton() { Font = new Font("Microsoft YaHei UI", 8.25f); }

        protected override void OnClick(EventArgs e)
        {
            if (Available) base.OnClick(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.Clear(BackColor);
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            var color = Available ? Color.FromArgb(244, 78, 72) : Color.FromArgb(205, 205, 205);
            var circle = new Rectangle(1, (Height - 12) / 2, 11, 11);
            using (var pen = new Pen(color, 1f)) e.Graphics.DrawEllipse(pen, circle);
            if (Checked)
                using (var brush = new SolidBrush(color)) e.Graphics.FillEllipse(brush, circle.X + 3, circle.Y + 3, 5, 5);
            TextRenderer.DrawText(e.Graphics, Text, Font, new Rectangle(17, 0, Width - 17, Height), Available ? ForeColor : Color.FromArgb(190, 190, 190), TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        }
    }

    internal sealed class AccentCheckBox : CheckBox
    {
        private bool _available = true;
        public bool Available { get => _available; set { if (_available == value) return; _available = value; TabStop = value; Invalidate(); } }

        public AccentCheckBox() { Font = new Font("Microsoft YaHei UI", 8.25f); }

        protected override void OnClick(EventArgs e)
        {
            if (Available) base.OnClick(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.Clear(BackColor);
            var color = Available ? Color.FromArgb(244, 78, 72) : Color.FromArgb(205, 205, 205);
            var box = new Rectangle(1, (Height - 12) / 2, 11, 11);
            using (var pen = new Pen(color, 1f)) e.Graphics.DrawRectangle(pen, box);
            if (Checked)
            {
                using (var brush = new SolidBrush(color)) e.Graphics.FillRectangle(brush, box);
                using (var pen = new Pen(Color.White, 1.4f))
                {
                    e.Graphics.DrawLine(pen, box.X + 2, box.Y + 6, box.X + 5, box.Y + 9);
                    e.Graphics.DrawLine(pen, box.X + 5, box.Y + 9, box.X + 10, box.Y + 2);
                }
            }
            TextRenderer.DrawText(e.Graphics, Text, Font, new Rectangle(17, 0, Width - 17, Height), Available ? ForeColor : Color.FromArgb(190, 190, 190), TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        }
    }
}
