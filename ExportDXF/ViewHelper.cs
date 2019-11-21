using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExportDXF
{
    internal static class ViewHelper
    {
        public static Bounds GetBounds(SolidWorks.Interop.sldworks.View view)
        {
            var outline = view.GetOutline() as double[];

            var minX = outline[0] / 0.0254;
            var minY = outline[1] / 0.0254;
            var maxX = outline[2] / 0.0254;
            var maxY = outline[3] / 0.0254;

            var width = Math.Abs(minX) + Math.Abs(maxX);
            var height = Math.Abs(minY) + Math.Abs(maxY);

            return new Bounds
            {
                X = minX,
                Y = minY,
                Width = width,
                Height = height
            };
        }

        public static Bend ClosestToBounds(Bounds bounds, IList<Bend> bends)
        {
            var hBends = bends.Where(b => GetAngleOrientation(b.BendLineAngle) == BendOrientation.Horizontal).ToList();
            var vBends = bends.Where(b => GetAngleOrientation(b.BendLineAngle) == BendOrientation.Vertical).ToList();

            Bend minVBend = null;
            double minVBendDist = double.MaxValue;

            foreach (var bend in vBends)
            {
                double distFromLft = Math.Abs(bend.X - bounds.Left);
                double distFromRgt = Math.Abs(bounds.Right - bend.X);

                double minDist = Math.Min(distFromLft, distFromRgt);

                if (minDist < minVBendDist)
                {
                    minVBendDist = minDist;
                    minVBend = bend;
                }
            }

            Bend minHBend = null;
            double minHBendDist = double.MaxValue;

            foreach (var bend in hBends)
            {
                double distFromBtm = Math.Abs(bend.Y - bounds.Bottom);
                double distFromTop = Math.Abs(bounds.Top - bend.Y);

                double minDist = Math.Min(distFromBtm, distFromTop);

                if (minDist < minHBendDist)
                {
                    minHBendDist = minDist;
                    minHBend = bend;
                }
            }

            return minHBendDist < minVBendDist ? minHBend : minVBend;
        }

        public static Bend SmallestYCoordinate(IList<Bend> bends)
        {
            double dist = double.MaxValue;
            int index = -1;

            for (int i = 0; i < bends.Count; i++)
            {
                var bend = bends[i];

                if (bend.Y < dist)
                {
                    dist = bend.Y;
                    index = i;
                }
            }

            return index == -1 ? null : bends[index];
        }

        public static Bend SmallestXCoordinate(IList<Bend> bends)
        {
            return bends.Min(b => b.X);
            double dist = double.MaxValue;
            int index = -1;

            for (int i = 0; i < bends.Count; i++)
            {
                var bend = bends[i];

                if (bend.X < dist)
                {
                    dist = bend.X;
                    index = i;
                }
            }

            return index == -1 ? null : bends[index];
        }

        public static BendDirection GetBendDirection(Note note)
        {
            var txt = note.GetText();

            return txt.ToUpper().Contains("UP") ? BendDirection.Up : BendDirection.Down;
        }

        public static IEnumerable<Note> GetBendNotes(SolidWorks.Interop.sldworks.View view)
        {
            return (view.GetNotes() as Array)?.Cast<Note>();
        }

        public static Note GetLeftMostNote(SolidWorks.Interop.sldworks.View view)
        {
            var notes = GetBendNotes(view);

            Note leftMostNote = null;
            var leftMostValue = double.MaxValue;

            foreach (var note in notes)
            {
                var pt = (note.GetTextPoint() as double[]);
                var x = pt[0];

                if (x < leftMostValue)
                {
                    leftMostValue = x;
                    leftMostNote = note;
                }
            }

            return leftMostNote;
        }

        public static Note GetBottomMostNote(SolidWorks.Interop.sldworks.View view)
        {
            var notes = GetBendNotes(view);

            Note btmMostNote = null;
            var btmMostValue = double.MaxValue;

            foreach (var note in notes)
            {
                var pt = (note.GetTextPoint() as double[]);
                var y = pt[1];

                if (y < btmMostValue)
                {
                    btmMostValue = y;
                    btmMostNote = note;
                }
            }

            return btmMostNote;
        }

        public static IEnumerable<double> GetBendAngles(SolidWorks.Interop.sldworks.View view)
        {
            var angles = new List<double>();
            var notes = GetBendNotes(view);

            foreach (var note in notes)
            {
                var angle = RadiansToDegrees(note.Angle);
                angles.Add(angle);
            }

            return angles;
        }

        public static List<Bend> GetBends(SolidWorks.Interop.sldworks.View view)
        {
            var bends = new List<Bend>();
            var notes = GetBendNotes(view);

            const string pattern = @"(?<DIRECTION>(UP|DOWN))\s*(?<ANGLE>(\d+(.\d*)?))°";

            foreach (var note in notes)
            {
                var pos = note.GetTextPoint2() as double[];

                var x = pos[0] / 0.0254;
                var y = pos[1] / 0.0254;

                var text = note.GetText();
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);

                if (!match.Success)
                    continue;

                var angle = double.Parse(match.Groups["ANGLE"].Value);
                var direection = match.Groups["DIRECTION"].Value;

                var bend = new Bend
                {
                    BendLineAngle = RadiansToDegrees(note.Angle),
                    Angle = angle,
                    Direction = direection == "UP" ? BendDirection.Up : BendDirection.Down,
                    X = x,
                    Y = y
                };

                bends.Add(bend);
            }

            return bends;
        }

        public static BendOrientation GetOrientation(SolidWorks.Interop.sldworks.View view)
        {
            var angles = GetBendAngles(view);

            var bends = GetBends(view);

            var vertical = 0;
            var horizontal = 0;

            foreach (var angle in angles)
            {
                var o = GetAngleOrientation(angle);

                switch (o)
                {
                    case BendOrientation.Horizontal:
                        horizontal++;
                        break;

                    case BendOrientation.Vertical:
                        vertical++;
                        break;
                }
            }

            if (vertical == 0 && horizontal == 0)
                return BendOrientation.Unknown;

            return vertical > horizontal ? BendOrientation.Vertical : BendOrientation.Horizontal;
        }

        public static BendOrientation GetAngleOrientation(double angleInDegrees)
        {
            if (angleInDegrees < 10 || angleInDegrees > 350)
                return BendOrientation.Horizontal;

            if (angleInDegrees > 170 && angleInDegrees < 190)
                return BendOrientation.Horizontal;

            if (angleInDegrees > 80 && angleInDegrees < 100)
                return BendOrientation.Vertical;

            if (angleInDegrees > 260 && angleInDegrees < 280)
                return BendOrientation.Vertical;

            return BendOrientation.Unknown;
        }

        public static double RadiansToDegrees(double angleInRadians)
        {
            return Math.Round(angleInRadians * 180.0 / Math.PI, 8);
        }
    }
}