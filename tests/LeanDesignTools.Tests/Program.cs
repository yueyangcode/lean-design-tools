using System;
using System.Collections.Generic;
using LeanDesignTools;

internal static class Program
{
    [STAThread]
    private static int Main()
    {
        try
        {
            var horizontal = new List<Box>
            {
                new Box { Id = 1, Left = 0, Width = 10 },
                new Box { Id = 2, Left = 15, Width = 10 },
                new Box { Id = 3, Left = 40, Width = 10 }
            };
            Geometry.DistributeHorizontally(horizontal);
            AssertNear(horizontal[1].Left, 20, "horizontal distribution");

            var anchor = new Box { Left = 10, Top = 20, Width = 40, Height = 30 };
            var item = new Box { Width = 10, Height = 10 };
            AssertNear(Geometry.AlignCenter(anchor, item), 25, "horizontal center");
            AssertNear(Geometry.AlignMiddle(anchor, item), 30, "vertical center");
            AssertNear(Geometry.AlignRight(anchor, item), 40, "right alignment");

            var moving = new Box { Left = 50, Top = 10, Width = 10, Height = 10 };
            Geometry.Snap(moving, new[] { new Box { Left = 20, Top = 12, Width = 20, Height = 5 } }, SnapDirection.Left, 100, 100);
            AssertNear(moving.Left, 40, "snap left to nearest overlapping shape");
            ShapeOperationSmokeTests.Run();
            Console.WriteLine("Geometry smoke tests passed.");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.GetType().FullName);
            Console.Error.WriteLine(ex.Message);
            Console.Error.WriteLine(ex.StackTrace);
            return 1;
        }
    }

    private static void AssertNear(float actual, float expected, string name)
    {
        if (Math.Abs(actual - expected) > 0.001f)
            throw new Exception($"{name}: expected {expected}, got {actual}");
    }
}
