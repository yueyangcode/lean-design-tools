using System;
using System.Collections.Generic;
using System.Linq;

namespace LeanDesignTools
{
    public enum SnapDirection { Left, Right, Up, Down }

    public sealed class Box
    {
        public int Id;
        public float Left;
        public float Top;
        public float Width;
        public float Height;
    }

    public static class Geometry
    {
        public static void DistributeHorizontally(IList<Box> boxes)
        {
            if (boxes == null || boxes.Count < 3) return;
            var ordered = boxes.OrderBy(x => x.Left).ToList();
            var span = ordered.Last().Left + ordered.Last().Width - ordered.First().Left;
            var gap = (span - ordered.Sum(x => x.Width)) / (ordered.Count - 1);
            var cursor = ordered.First().Left + ordered.First().Width + gap;
            for (var i = 1; i < ordered.Count - 1; i++)
            {
                ordered[i].Left = cursor;
                cursor += ordered[i].Width + gap;
            }
        }

        public static void DistributeVertically(IList<Box> boxes)
        {
            if (boxes == null || boxes.Count < 3) return;
            var ordered = boxes.OrderBy(x => x.Top).ToList();
            var span = ordered.Last().Top + ordered.Last().Height - ordered.First().Top;
            var gap = (span - ordered.Sum(x => x.Height)) / (ordered.Count - 1);
            var cursor = ordered.First().Top + ordered.First().Height + gap;
            for (var i = 1; i < ordered.Count - 1; i++)
            {
                ordered[i].Top = cursor;
                cursor += ordered[i].Height + gap;
            }
        }

        public static float AlignLeft(Box anchor) => anchor.Left;
        public static float AlignRight(Box anchor, Box item) => anchor.Left + anchor.Width - item.Width;
        public static float AlignCenter(Box anchor, Box item) => anchor.Left + (anchor.Width - item.Width) / 2f;
        public static float AlignTop(Box anchor) => anchor.Top;
        public static float AlignBottom(Box anchor, Box item) => anchor.Top + anchor.Height - item.Height;
        public static float AlignMiddle(Box anchor, Box item) => anchor.Top + (anchor.Height - item.Height) / 2f;

        public static void Snap(Box moving, IEnumerable<Box> others, SnapDirection direction, float slideWidth, float slideHeight)
        {
            var candidates = others ?? Enumerable.Empty<Box>();
            if (direction == SnapDirection.Left)
            {
                var edge = candidates.Where(x => Overlaps(moving.Top, moving.Height, x.Top, x.Height) && x.Left + x.Width <= moving.Left)
                    .Select(x => x.Left + x.Width).DefaultIfEmpty(0).Max();
                moving.Left = edge;
            }
            else if (direction == SnapDirection.Right)
            {
                var edge = candidates.Where(x => Overlaps(moving.Top, moving.Height, x.Top, x.Height) && x.Left >= moving.Left + moving.Width)
                    .Select(x => x.Left).DefaultIfEmpty(slideWidth).Min();
                moving.Left = edge - moving.Width;
            }
            else if (direction == SnapDirection.Up)
            {
                var edge = candidates.Where(x => Overlaps(moving.Left, moving.Width, x.Left, x.Width) && x.Top + x.Height <= moving.Top)
                    .Select(x => x.Top + x.Height).DefaultIfEmpty(0).Max();
                moving.Top = edge;
            }
            else
            {
                var edge = candidates.Where(x => Overlaps(moving.Left, moving.Width, x.Left, x.Width) && x.Top >= moving.Top + moving.Height)
                    .Select(x => x.Top).DefaultIfEmpty(slideHeight).Min();
                moving.Top = edge - moving.Height;
            }
        }

        private static bool Overlaps(float startA, float lengthA, float startB, float lengthB) =>
            startA < startB + lengthB && startB < startA + lengthA;
    }
}
