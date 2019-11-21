using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ExportDXF
{
    public static class SolidWorksExtensions
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

            return null;
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
}