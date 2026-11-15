using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using LeanDesignTools;

internal static class ShapeOperationSmokeTests
{
    public static void Run()
    {
        TestAlignmentAndDistribution();
        TestSizeDockRotateAndOrdering();
        TestSelectionVectorAndClipboard();
        TestSnapTextAndMargins();
        TestSplitAndCombine();
        Console.WriteLine("Shape operation smoke tests passed.");
    }

    private static Fixture NewFixture()
    {
        var app = new FakeApp();
        var a = app.Slide.Shapes.Add(new FakeShape(1, 10, 10, 30, 20));
        var b = app.Slide.Shapes.Add(new FakeShape(2, 50, 35, 10, 8));
        var c = app.Slide.Shapes.Add(new FakeShape(3, 80, 60, 20, 10));
        app.Select(a, b, c);
        typeof(DesignPane).GetField("PowerPoint", BindingFlags.Static | BindingFlags.NonPublic).SetValue(null, app);
        return new Fixture(app, a, b, c);
    }

    private static void TestAlignmentAndDistribution()
    {
        var f = NewFixture();
        ShapeOperations.Align(AlignMode.Left); Equal(f.B.Left, 10, "align left");
        f.B.Left = 50; ShapeOperations.Align(AlignMode.Center); Equal(f.B.Left, 20, "align center");
        f.B.Left = 50; ShapeOperations.Align(AlignMode.Right); Equal(f.B.Left, 30, "align right");
        ShapeOperations.Align(AlignMode.Top); Equal(f.B.Top, 10, "align top");
        f.B.Top = 35; ShapeOperations.Align(AlignMode.Middle); Equal(f.B.Top, 16, "align middle");
        f.B.Top = 35; ShapeOperations.Align(AlignMode.Bottom); Equal(f.B.Top, 22, "align bottom");

        f.A.Left = 0; f.A.Width = 10; f.B.Left = 15; f.B.Width = 10; f.C.Left = 40; f.C.Width = 10;
        ShapeOperations.Distribute(true); Equal(f.B.Left, 20, "distribute horizontal");
        f.A.Top = 0; f.A.Height = 10; f.B.Top = 15; f.B.Height = 10; f.C.Top = 40; f.C.Height = 10;
        ShapeOperations.Distribute(false); Equal(f.B.Top, 20, "distribute vertical");
    }

    private static void TestSizeDockRotateAndOrdering()
    {
        var f = NewFixture();
        ShapeOperations.MatchSize(true, false); Equal(f.B.Width, 30, "same width"); Equal(f.B.Height, 8, "same width keeps height");
        f.B.Width = 10; ShapeOperations.MatchSize(false, true); Equal(f.B.Height, 20, "same height"); Equal(f.B.Width, 10, "same height keeps width");
        ShapeOperations.MatchSize(true, true); Equal(f.B.Width, 30, "same size width"); Equal(f.B.Height, 20, "same size height");

        f.App.Select(f.B); f.B.Width = 10; f.B.Height = 8;
        ShapeOperations.SetDimensions(2m, null); Equal(f.B.Width, 56.6929f, "set width in centimeters", 0.001f); Equal(f.B.Height, 8, "set width keeps height");
        ShapeOperations.SetDimensions(null, 1.5m); Equal(f.B.Height, 42.5197f, "set height in centimeters", 0.001f); Equal(f.B.Width, 56.6929f, "set height keeps width", 0.001f);

        f.B.Width = 10; f.B.Height = 8; f.App.Select(f.B);
        ShapeOperations.Dock(-1, 2); Equal(f.B.Left, 0, "dock left");
        ShapeOperations.Dock(0, 2); Equal(f.B.Left, 45, "dock horizontal center");
        ShapeOperations.Dock(1, 2); Equal(f.B.Left, 90, "dock right");
        ShapeOperations.Dock(2, -1); Equal(f.B.Top, 0, "dock top");
        ShapeOperations.Dock(2, 0); Equal(f.B.Top, 36, "dock vertical center");
        ShapeOperations.Dock(2, 1); Equal(f.B.Top, 72, "dock bottom");
        ShapeOperations.Dock(0, 0); Equal(f.B.Left, 45, "slide center x"); Equal(f.B.Top, 36, "slide center y");
        ShapeOperations.Stretch(true); Equal(f.B.Left, 0, "stretch horizontal left"); Equal(f.B.Width, 100, "stretch horizontal width");
        ShapeOperations.Stretch(false); Equal(f.B.Top, 0, "stretch vertical top"); Equal(f.B.Height, 80, "stretch vertical height");

        ShapeOperations.Rotate(90); Equal(f.B.Rotation, 90, "rotate right");
        ShapeOperations.Rotate(-90); Equal(f.B.Rotation, 0, "rotate left");
        ShapeOperations.Flip(true); Equal(f.B.LastFlip, 0, "flip horizontal");
        ShapeOperations.Flip(false); Equal(f.B.LastFlip, 1, "flip vertical");
        foreach (var command in new[] { 0, 2, 3, 1 }) { ShapeOperations.ZOrder(command); Equal(f.B.LastZOrder, command, "z-order " + command); }

        f.App.Select(f.A, f.B); var groupedRange = f.App.Selection.ShapeRange; ShapeOperations.Group(); True(groupedRange.GroupCalled, "group");
        f.App.Select(f.A); var ungroupedRange = f.App.Selection.ShapeRange; ShapeOperations.Ungroup(); True(ungroupedRange.UngroupCalled, "ungroup");
    }

    private static void TestSelectionVectorAndClipboard()
    {
        var f = NewFixture();
        ShapeOperations.SelectAll(); True(f.App.Slide.Shapes.SelectedAll, "select all");
        foreach (var id in new[] { "SelectionPane", "ShapesUnion", "ShapesCombine", "ShapesFragment", "ShapesIntersect", "ShapesSubtract" })
        {
            ShapeOperations.ExecuteMso(id, id == "SelectionPane" ? 0 : 2);
            True(f.App.CommandBars.Executed.Contains(id), "ExecuteMso " + id);
        }

        f.App.Select(f.A); ShapeOperations.CopyGeometry(true, false, false);
        ShapeOperations.CopyGeometry(false, true, false);
        ShapeOperations.CopyGeometry(false, false, true);
        f.App.Select(f.B); ShapeOperations.PasteGeometry(true, false, false); Equal(f.B.Left, 10, "paste position left"); Equal(f.B.Top, 10, "paste position top");
        f.B.Width = 5; f.B.Height = 6; ShapeOperations.PasteGeometry(false, true, false); Equal(f.B.Width, 30, "paste width"); Equal(f.B.Height, 6, "paste width keeps height");
        ShapeOperations.PasteGeometry(false, false, true); Equal(f.B.Width, 30, "paste height keeps width"); Equal(f.B.Height, 20, "paste height");
        f.B.Left = 70; f.B.Top = 70; f.B.Width = 5; f.B.Height = 5; ShapeOperations.PasteGeometry(true, true, true);
        Equal(f.B.Left, 10, "paste all left"); Equal(f.B.Top, 10, "paste all top"); Equal(f.B.Width, 30, "paste all width"); Equal(f.B.Height, 20, "paste all height");
    }

    private static void TestSnapTextAndMargins()
    {
        var f = NewFixture();
        f.A.Left = 40; f.A.Top = 20; f.A.Width = 10; f.A.Height = 10;
        f.B.Left = 20; f.B.Top = 22; f.B.Width = 15; f.B.Height = 5;
        f.C.Left = 60; f.C.Top = 22; f.C.Width = 10; f.C.Height = 5;
        f.App.Select(f.A); ShapeOperations.Snap(SnapDirection.Left); Equal(f.A.Left, 35, "snap left");
        f.A.Left = 40; ShapeOperations.Snap(SnapDirection.Right); Equal(f.A.Left, 50, "snap right");
        f.B.Left = 42; f.B.Top = 5; f.B.Width = 5; f.B.Height = 10; f.C.Left = 42; f.C.Top = 45; f.C.Width = 5; f.C.Height = 10;
        f.A.Left = 40; f.A.Top = 20; ShapeOperations.Snap(SnapDirection.Up); Equal(f.A.Top, 15, "snap up");
        f.A.Left = 40; f.A.Top = 20; ShapeOperations.Snap(SnapDirection.Down); Equal(f.A.Top, 35, "snap down");

        ShapeOperations.SetTextBoxMode(0); Equal(f.A.TextFrame2.AutoSize, 0, "text free mode");
        ShapeOperations.SetTextBoxMode(1); Equal(f.A.TextFrame2.AutoSize, 1, "text fit shape mode");
        ShapeOperations.SetTextBoxMode(2); Equal(f.A.TextFrame2.AutoSize, 2, "text shrink mode");
        ShapeOperations.SetWordWrap(true); Equal(f.A.TextFrame2.WordWrap, -1, "word wrap on");
        ShapeOperations.SetWordWrap(false); Equal(f.A.TextFrame2.WordWrap, 0, "word wrap off");
        ShapeOperations.SetMargins(0.25m, 0.5m, 0.13m, 0.2m);
        Equal(f.A.TextFrame.MarginLeft, 7.0866f, "margin left", 0.001f);
        Equal(f.A.TextFrame.MarginRight, 14.1732f, "margin right", 0.001f);
        Equal(f.A.TextFrame.MarginTop, 3.6850f, "margin top", 0.001f);
        Equal(f.A.TextFrame.MarginBottom, 5.6693f, "margin bottom", 0.001f);
    }

    private static void TestSplitAndCombine()
    {
        var f = NewFixture();
        f.App.Select(f.A); f.A.TextFrame.TextRange.Text = "A\rB\vC";
        ShapeOperations.SplitText(true);
        var paragraphParts = f.App.Slide.Shapes.Items.Where(x => x.Id >= 100).ToList();
        Equal(paragraphParts.Count, 2, "paragraph split count");

        f = NewFixture(); f.App.Select(f.A); f.A.TextFrame.TextRange.Text = "AB\rC";
        ShapeOperations.SplitText(false);
        var lineParts = f.App.Slide.Shapes.Items.Where(x => x.Id >= 100).OrderBy(x => x.Top).ToList();
        Equal(lineParts.Count, 3, "character split count");
        f.App.Select(lineParts.ToArray()); ShapeOperations.CombineText(true);
        Equal(lineParts[0].TextFrame.TextRange.Text, "A\rB\rC", "combine with line breaks");

        f = NewFixture(); f.A.TextFrame.TextRange.Text = "A"; f.B.TextFrame.TextRange.Text = "B"; f.App.Select(f.A, f.B);
        ShapeOperations.CombineText(false); Equal(f.A.TextFrame.TextRange.Text, "AB", "combine characters");
        True(!f.App.Slide.Shapes.Items.Contains(f.B), "combine deletes remaining shapes");
    }

    private static void Equal(float actual, float expected, string name, float tolerance = 0.0001f)
    {
        if (Math.Abs(actual - expected) > tolerance) throw new Exception($"{name}: expected {expected}, got {actual}");
    }

    private static void Equal(int actual, int expected, string name) { if (actual != expected) throw new Exception($"{name}: expected {expected}, got {actual}"); }
    private static void Equal(string actual, string expected, string name) { if (actual != expected) throw new Exception($"{name}: expected '{expected}', got '{actual}'"); }
    private static void True(bool value, string name) { if (!value) throw new Exception(name + " failed"); }

    private sealed class Fixture
    {
        public readonly FakeApp App; public readonly FakeShape A; public readonly FakeShape B; public readonly FakeShape C;
        public Fixture(FakeApp app, FakeShape a, FakeShape b, FakeShape c) { App = app; A = a; B = b; C = c; }
    }
}

public sealed class FakeApp
{
    public FakeWindow ActiveWindow { get; } = new FakeWindow();
    public FakePresentation ActivePresentation { get; } = new FakePresentation();
    public FakeCommandBars CommandBars { get; } = new FakeCommandBars();
    public FakeSlide Slide => ActiveWindow.View.Slide;
    public FakeSelection Selection => ActiveWindow.Selection;
    public int UndoEntries { get; private set; }
    public void StartNewUndoEntry() { UndoEntries++; }
    public void Select(params FakeShape[] shapes) { ActiveWindow.Selection.ShapeRange = new FakeRange(shapes); }
}

public sealed class FakeWindow
{
    public FakeSelection Selection { get; } = new FakeSelection();
    public FakeView View { get; } = new FakeView();
}

public sealed class FakeSelection { public int Type { get; set; } = 2; public FakeRange ShapeRange { get; set; } = new FakeRange(); }
public sealed class FakeView { public FakeSlide Slide { get; } = new FakeSlide(); }
public sealed class FakePresentation { public FakePageSetup PageSetup { get; } = new FakePageSetup(); }
public sealed class FakePageSetup { public float SlideWidth { get; set; } = 100; public float SlideHeight { get; set; } = 80; }
public sealed class FakeCommandBars { public List<string> Executed { get; } = new List<string>(); public void ExecuteMso(string id) { Executed.Add(id); } }
public sealed class FakeSlide { public FakeShapes Shapes { get; } = new FakeShapes(); }

public sealed class FakeShapes
{
    private int _nextId = 100;
    public List<FakeShape> Items { get; } = new List<FakeShape>();
    public int Count => Items.Count;
    public bool SelectedAll { get; private set; }
    public FakeShape Add(FakeShape shape) { shape.Owner = this; Items.Add(shape); return shape; }
    public FakeShape AddClone(FakeShape source) { var clone = source.Clone(_nextId++); return Add(clone); }
    public FakeShape Item(int index) => Items[index - 1];
    public FakeRange Range(object _) => new FakeRange(Items.ToArray(), () => SelectedAll = true);
    public void Remove(FakeShape shape) { Items.Remove(shape); }
}

public sealed class FakeRange
{
    private readonly List<FakeShape> _items;
    private readonly Action _select;
    public FakeRange(params FakeShape[] items) : this(items, null) { }
    public FakeRange(FakeShape[] items, Action select) { _items = items.ToList(); _select = select; }
    public int Count => _items.Count;
    public bool GroupCalled { get; private set; }
    public bool UngroupCalled { get; private set; }
    public FakeShape Item(int index) => _items[index - 1];
    public object Group() { GroupCalled = true; return new object(); }
    public object Ungroup() { UngroupCalled = true; return new object(); }
    public void Select() { _select?.Invoke(); }
}

public sealed class FakeShape
{
    public int Id { get; }
    public float Left { get; set; }
    public float Top { get; set; }
    public float Width { get; set; }
    public float Height { get; set; }
    public float Rotation { get; set; }
    public int HasTextFrame { get; set; } = -1;
    public int LastFlip { get; private set; } = -1;
    public int LastZOrder { get; private set; } = -1;
    public FakeTextFrame TextFrame { get; } = new FakeTextFrame();
    public FakeTextFrame2 TextFrame2 { get; } = new FakeTextFrame2();
    public FakeShapes Owner { get; set; }
    public FakeShape(int id, float left, float top, float width, float height) { Id = id; Left = left; Top = top; Width = width; Height = height; }
    public void Flip(int command) { LastFlip = command; }
    public void ZOrder(int command) { LastZOrder = command; }
    public FakeRange Duplicate() => new FakeRange(Owner.AddClone(this));
    public void Delete() { Owner?.Remove(this); }
    public FakeShape Clone(int id)
    {
        var clone = new FakeShape(id, Left, Top, Width, Height) { Rotation = Rotation, HasTextFrame = HasTextFrame };
        clone.TextFrame.TextRange.Text = TextFrame.TextRange.Text;
        clone.TextFrame2.AutoSize = TextFrame2.AutoSize; clone.TextFrame2.WordWrap = TextFrame2.WordWrap;
        return clone;
    }
}

public sealed class FakeTextFrame
{
    public FakeTextRange TextRange { get; } = new FakeTextRange();
    public int AutoSize { get; set; }
    public float MarginLeft { get; set; }
    public float MarginRight { get; set; }
    public float MarginTop { get; set; }
    public float MarginBottom { get; set; }
}

public sealed class FakeTextFrame2 { public int AutoSize { get; set; } public int WordWrap { get; set; } }
public sealed class FakeTextRange { public string Text { get; set; } = ""; }
