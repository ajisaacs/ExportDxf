using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExportDXF
{
    public static class Extensions
    {
        public static void AppendText(this RichTextBox box, string text, Color color)
        {
            box.SelectionStart = box.TextLength;
            box.SelectionLength = 0;

            box.SelectionColor = color;
            box.AppendText(text);
            box.SelectionColor = box.ForeColor;
        }

        public static string ToReadableFormat(this TimeSpan ts)
        {
            var s = new StringBuilder();

            if (ts.TotalHours >= 1)
            {
                var hrs = ts.Hours + ts.Days * 24.0;

                s.Append(string.Format("{0}hrs ", hrs));
                s.Append(string.Format("{0}min ", ts.Minutes));
                s.Append(string.Format("{0}sec", ts.Seconds));
            }
            else if (ts.TotalMinutes >= 1)
            {
                s.Append(string.Format("{0}min ", ts.Minutes));
                s.Append(string.Format("{0}sec", ts.Seconds));
            }
            else
            {
                s.Append(string.Format("{0} seconds", ts.Seconds));
            }

            return s.ToString();
        }

        public static string PunctuateList(this IEnumerable<string> stringList)
        {
            var list = stringList.ToList();

            switch (list.Count)
            {
                case 0:
                    return string.Empty;

                case 1:
                    return list[0];

                case 2:
                    return string.Format("{0} and {1}", list[0], list[1]);

                default:
                    var s = string.Empty;

                    for (int i = 0; i < list.Count - 1; i++)
                        s += list[i] + ", ";

                    s += "and " + list.Last();

                    return s;
            }
        }
    }
}