using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
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

        public static List<BomTableAnnotation> GetBomTables(this DrawingDoc drawing)
        {
            var model = drawing as ModelDoc2;

            return model.GetAllFeaturesByTypeName("BomFeat")
                .Select(f => f.GetSpecificFeature2() as BomFeature)
                .Select(f => (f.GetTableAnnotations() as Array)?.Cast<BomTableAnnotation>().FirstOrDefault())
                .ToList();
        }

        public static void SetLightweightToResolved(this Component2 component)
        {
            var suppressionState = (swComponentSuppressionState_e)component.GetSuppression();

            switch (suppressionState)
            {
                case swComponentSuppressionState_e.swComponentFullyResolved:
                case swComponentSuppressionState_e.swComponentResolved:
                    return;

                case swComponentSuppressionState_e.swComponentFullyLightweight:
                case swComponentSuppressionState_e.swComponentLightweight:
                    var error = (swSuppressionError_e)component.SetSuppression2((int)swComponentSuppressionState_e.swComponentResolved);
                    break;
            }
        }
    }
}