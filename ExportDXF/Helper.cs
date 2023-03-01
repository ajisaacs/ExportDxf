using SolidWorks.Interop.sldworks;
using System.Diagnostics;

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
            flatPatternFeatureData.SimplifyBends = false;

            var ret = flatPattern.ModifyDefinition(flatPatternFeatureData, model, null);
        }
    }
}