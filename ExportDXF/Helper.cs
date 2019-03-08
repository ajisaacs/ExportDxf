using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExportDXF
{
    public static class Helper
    {
        public static Feature GetFeatureByTypeName(this ModelDoc2 model, string featureName)
        {
            var feature = model.FirstFeature() as Feature;

            while (feature != null)
            {
                if (feature.GetTypeName() == featureName)
                    return feature;

                feature = feature.GetNextFeature() as Feature;
            }

            return feature;
        }

        public static List<Feature> GetAllFeaturesByTypeName(this ModelDoc2 model, string featureName)
        {
            var feature = model.FirstFeature() as Feature;
            var list = new List<Feature>();

            while (feature != null)
            {
                var name = feature.GetTypeName();

                if (name == featureName)
                    list.Add(feature);

                feature = feature.GetNextFeature() as Feature;
            }

            return list;
        }

        public static List<Feature> GetAllSubFeaturesByTypeName(this Feature feature, string subFeatureName)
        {
            var subFeature = feature.GetFirstSubFeature() as Feature;
            var list = new List<Feature>();

            while (subFeature != null)
            {
                Debug.WriteLine(subFeature.GetTypeName2());
                if (subFeature.GetTypeName() == subFeatureName)
                    list.Add(subFeature);

                subFeature = subFeature.GetNextSubFeature() as Feature;
            }

            return list;
        }

        public static bool HasFlatPattern(this ModelDoc2 model)
        {
            return model.GetBendState() != (int)swSMBendState_e.swSMBendStateNone;
        }

        public static bool IsSheetMetal(this ModelDoc2 model)
        {
            if (model is PartDoc == false)
                return false;

            if (model.HasFlatPattern() == false)
                return false;

            return true;
        }

        public static bool IsPart(this ModelDoc2 model)
        {
            return model is PartDoc;
        }

        public static bool IsDrawing(this ModelDoc2 model)
        {
            return model is DrawingDoc;
        }

        public static bool IsAssembly(this ModelDoc2 model)
        {
            return model is AssemblyDoc;
        }

        public static string GetTitle(this Component2 component)
        {
            var model = component.GetModelDoc2() as ModelDoc2;
            return model.GetTitle();
        }

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

        public static string GetNumWithSuffix(int i)
        {
            if (i >= 11 && i <= 13)
                return i.ToString() + "th";

            var j = i % 10;

            switch (j)
            {
                case 1: return i.ToString() + "st";
                case 2: return i.ToString() + "nd";
                case 3: return i.ToString() + "rd";
                default: return i.ToString() + "th";
            }
        }

        public static int IndexOfColumnType(this TableAnnotation table, swTableColumnTypes_e columnType)
        {
            for (int columnIndex = 0; columnIndex < table.ColumnCount; ++columnIndex)
            {
                var currentColumnType = (swTableColumnTypes_e)table.GetColumnType(columnIndex);

                if (currentColumnType == columnType)
                    return columnIndex;
            }

            return -1;
        }

        public static int IndexOfColumnTitle(this TableAnnotation table, string columnTitle)
        {
            var lowercaseColumnTitle = columnTitle.ToLower();

            for (int columnIndex = 0; columnIndex < table.ColumnCount; ++columnIndex)
            {
                var currentColumnType = table.GetColumnTitle(columnIndex);

                if (currentColumnType.ToLower() == lowercaseColumnTitle)
                    return columnIndex;
            }

            return -1;
        }

        public static Dimension GetDimension(this Feature feature, string dimName)
        {
            return feature?.Parameter(dimName) as Dimension;
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

    public static class Units
    {
        /// <summary>
        /// Multiply factor needed to convert the desired units to meters.
        /// </summary>
        public static double ScaleFactor = 0.0254; // inches to meters

        public static double ToSldWorks(this double d)
        {
            return Math.Round(d * ScaleFactor, 8);
        }

        public static double FromSldWorks(this double d)
        {
            return Math.Round(d / ScaleFactor, 8);
        }
    }
}