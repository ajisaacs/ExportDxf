using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace ExportDXF
{
    public static class Helper
    {
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

        public static void UncheckFlatPatternCornerTreatment(ModelDoc2 model)
        {
            var flatPattern = model.GetFeatureByTypeName("FlatPattern");
            var flatPatternFeatureData = flatPattern.GetDefinition() as FlatPatternFeatureData;

            flatPatternFeatureData.CornerTreatment = false;
            flatPatternFeatureData.SimplifyBends = true;

            var ret = flatPattern.ModifyDefinition(flatPatternFeatureData, model, null);
        }

        public static bool SetFlatPatternSuppressionState(ModelDoc2 model, swComponentSuppressionState_e suppressionState)
        {
            

            var flatPattern = model.GetFeatureByTypeName("FlatPattern");

            flatPattern.SetSuppression((int)suppressionState);

            return flatPattern.IsSuppressed();
        }

        public static string RemoveXmlTags(string input)
        {
            // Define the regular expression pattern to match XML tags
            string pattern = @"<[^>]+>";

            // Replace all matches of the pattern with an empty string
            string result = Regex.Replace(input, pattern, "");

            return result;
        }


    }
}