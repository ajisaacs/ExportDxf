using ExportDXF.Extensions;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExportDXF.Utilities
{
    /// <summary>
    /// Utility class for SolidWorks-specific operations and helper functions.
    /// </summary>
    public static class SolidWorksHelper
    {
        #region Feature Names Constants

        private const string FLAT_PATTERN_FEATURE = "FlatPattern";
        private const string SHEET_METAL_FEATURE = "SheetMetal";
        private const string BASE_FLANGE_FEATURE = "BaseFlange";
        private const string FLAT_PATTERN_FOLDER = "Flat-Pattern";

        #endregion

        #region Sheet Metal Operations

        /// <summary>
        /// Gets sheet metal properties from the model.
        /// </summary>
        /// <param name="model">The model to extract properties from.</param>
        /// <returns>Sheet metal properties, or null if not a sheet metal part.</returns>
        public static SheetMetalProperties GetSheetMetalProperties(ModelDoc2 model)
        {
            if (model == null)
                throw new ArgumentNullException(nameof(model));

            var sheetMetalFeature = model.GetFeatureByTypeName(SHEET_METAL_FEATURE);

            if (sheetMetalFeature == null)
                return null;

            var sheetMetalData = sheetMetalFeature.GetDefinition() as SheetMetalFeatureData;

            if (sheetMetalData == null)
                return null;

            return new SheetMetalProperties
            {
                Thickness = sheetMetalData.Thickness.FromSolidWorksToInches(),
                KFactor = sheetMetalData.KFactor,
                BendRadius = sheetMetalData.BendRadius.FromSolidWorksToInches(),
                BendAllowance = sheetMetalData.BendAllowance.FromSolidWorksToInches(),
                AutoRelief = sheetMetalData.UseAutoRelief,
                ReliefRatio = sheetMetalData.ReliefRatio
            };
        }

        /// <summary>
        /// Configures flat pattern settings for optimal DXF export.
        /// Unchecks corner treatment and enables simplify bends.
        /// </summary>
        /// <param name="model">The model containing the flat pattern.</param>
        /// <returns>True if settings were successfully modified.</returns>
        public static bool ConfigureFlatPatternSettings(ModelDoc2 model)
        {
            if (model == null)
                throw new ArgumentNullException(nameof(model));

            var flatPattern = model.GetFeatureByTypeName(FLAT_PATTERN_FEATURE);

            if (flatPattern == null)
                return false;

            var featureData = flatPattern.GetDefinition() as FlatPatternFeatureData;

            if (featureData == null)
                return false;

            try
            {
                // Configure for cleaner DXF output
                featureData.CornerTreatment = false;  // Remove corner treatments
                featureData.SimplifyBends = true;     // Simplify bend representations

                return flatPattern.ModifyDefinition(featureData, model, null);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Sets the suppression state of the flat pattern feature.
        /// </summary>
        /// <param name="model">The model containing the flat pattern.</param>
        /// <param name="suppressionState">The desired suppression state.</param>
        /// <returns>True if the flat pattern is suppressed after the operation.</returns>
        public static bool SetFlatPatternSuppressionState(
            ModelDoc2 model,
            swComponentSuppressionState_e suppressionState)
        {
            if (model == null)
                throw new ArgumentNullException(nameof(model));

            var flatPattern = model.GetFeatureByTypeName(FLAT_PATTERN_FEATURE);

            if (flatPattern == null)
                return false;

            try
            {
                flatPattern.SetSuppression((int)suppressionState);
                return flatPattern.IsSuppressed();
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Gets the flat pattern feature from a model.
        /// </summary>
        /// <param name="model">The model to search.</param>
        /// <returns>The flat pattern feature, or null if not found.</returns>
        public static Feature GetFlatPatternFeature(ModelDoc2 model)
        {
            return model?.GetFeatureByTypeName(FLAT_PATTERN_FEATURE);
        }

        /// <summary>
        /// Checks if a model has a flat pattern feature.
        /// </summary>
        /// <param name="model">The model to check.</param>
        /// <returns>True if the model has a flat pattern feature.</returns>
        public static bool HasFlatPattern(ModelDoc2 model)
        {
            return GetFlatPatternFeature(model) != null;
        }

        #endregion

        #region Feature Operations

        /// <summary>
        /// Gets all features of a specific type from a model.
        /// </summary>
        /// <param name="model">The model to search.</param>
        /// <param name="featureTypeName">The name of the feature type to find.</param>
        /// <returns>A list of features of the specified type.</returns>
        public static List<Feature> GetFeaturesByTypeName(ModelDoc2 model, string featureTypeName)
        {
            if (model == null || string.IsNullOrEmpty(featureTypeName))
                return new List<Feature>();

            var features = new List<Feature>();
            var feature = model.FirstFeature() as Feature;

            while (feature != null)
            {
                if (string.Equals(feature.GetTypeName2(), featureTypeName, StringComparison.OrdinalIgnoreCase))
                {
                    features.Add(feature);
                }

                // Check sub-features
                var subFeature = feature.GetFirstSubFeature() as Feature;
                while (subFeature != null)
                {
                    if (string.Equals(subFeature.GetTypeName2(), featureTypeName, StringComparison.OrdinalIgnoreCase))
                    {
                        features.Add(subFeature);
                    }
                    subFeature = subFeature.GetNextSubFeature() as Feature;
                }

                feature = feature.GetNextFeature() as Feature;
            }

            return features;
        }

        /// <summary>
        /// Suppresses or unsuppresses a feature by name.
        /// </summary>
        /// <param name="model">The model containing the feature.</param>
        /// <param name="featureName">The name of the feature to suppress/unsuppress.</param>
        /// <param name="suppress">True to suppress, false to unsuppress.</param>
        /// <returns>True if the operation was successful.</returns>
        public static bool SetFeatureSuppressionByName(ModelDoc2 model, string featureName, bool suppress)
        {
            if (model == null || string.IsNullOrEmpty(featureName))
                return false;

            try
            {
                // Use Extension.SelectByID2 to find and select the feature
                var modelExtension = model.Extension;
                bool selected = modelExtension.SelectByID2(featureName, "BODYFEATURE", 0, 0, 0, false, 0, null, 0);

                if (!selected)
                    return false;

                var selectionManager = model.SelectionManager as SelectionMgr;
                var feature = selectionManager.GetSelectedObject6(1, -1) as Feature;

                if (feature == null)
                    return false;

                var suppressionState = suppress
                    ? swFeatureSuppressionAction_e.swSuppressFeature
                    : swFeatureSuppressionAction_e.swUnSuppressFeature;

                feature.SetSuppression((int)suppressionState);

                // Clear the selection
                model.ClearSelection2(true);

                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Configuration Operations

        /// <summary>
        /// Gets all configuration names from a model.
        /// </summary>
        /// <param name="model">The model to query.</param>
        /// <returns>A list of configuration names.</returns>
        public static List<string> GetConfigurationNames(ModelDoc2 model)
        {
            if (model == null)
                return new List<string>();

            var configNames = (string[])model.GetConfigurationNames();
            return configNames?.ToList() ?? new List<string>();
        }

        /// <summary>
        /// Activates a specific configuration.
        /// </summary>
        /// <param name="model">The model to modify.</param>
        /// <param name="configurationName">The name of the configuration to activate.</param>
        /// <returns>True if the configuration was successfully activated.</returns>
        public static bool ActivateConfiguration(ModelDoc2 model, string configurationName)
        {
            if (model == null || string.IsNullOrEmpty(configurationName))
                return false;

            try
            {
                return model.ShowConfiguration2(configurationName);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Gets the active configuration name.
        /// </summary>
        /// <param name="model">The model to query.</param>
        /// <returns>The name of the active configuration, or null if unavailable.</returns>
        public static string GetActiveConfigurationName(ModelDoc2 model)
        {
            if (model == null)
                return null;

            try
            {
                var config = model.GetActiveConfiguration() as Configuration;
                return config?.Name;
            }
            catch
            {
                return null;
            }
        }

        #endregion

        #region Custom Properties

        /// <summary>
        /// Gets a custom property value from a model.
        /// </summary>
        /// <param name="model">The model to query.</param>
        /// <param name="propertyName">The name of the property.</param>
        /// <param name="configurationName">The configuration name (empty string for file-level properties).</param>
        /// <returns>The property value, or null if not found.</returns>
        public static string GetCustomProperty(ModelDoc2 model, string propertyName, string configurationName = "")
        {
            if (model == null || string.IsNullOrEmpty(propertyName))
                return null;

            try
            {
                var customPropertyManager = model.Extension.CustomPropertyManager[configurationName ?? ""];
                return customPropertyManager.Get(propertyName);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Sets a custom property value on a model.
        /// </summary>
        /// <param name="model">The model to modify.</param>
        /// <param name="propertyName">The name of the property.</param>
        /// <param name="propertyValue">The value to set.</param>
        /// <param name="configurationName">The configuration name (empty string for file-level properties).</param>
        /// <returns>True if the property was successfully set.</returns>
        public static bool SetCustomProperty(ModelDoc2 model, string propertyName, string propertyValue, string configurationName = "")
        {
            if (model == null || string.IsNullOrEmpty(propertyName))
                return false;

            try
            {
                var customPropertyManager = model.Extension.CustomPropertyManager[configurationName ?? ""];
                var result = customPropertyManager.Add3(propertyName, (int)swCustomInfoType_e.swCustomInfoText, propertyValue, (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Gets all custom property names from a model.
        /// </summary>
        /// <param name="model">The model to query.</param>
        /// <param name="configurationName">The configuration name (empty string for file-level properties).</param>
        /// <returns>A list of custom property names.</returns>
        public static List<string> GetCustomPropertyNames(ModelDoc2 model, string configurationName = "")
        {
            if (model == null)
                return new List<string>();

            try
            {
                var customPropertyManager = model.Extension.CustomPropertyManager[configurationName ?? ""];
                var propertyNames = (string[])customPropertyManager.GetNames();
                return propertyNames?.ToList() ?? new List<string>();
            }
            catch
            {
                return new List<string>();
            }
        }

        #endregion

        #region Material Operations

        /// <summary>
        /// Gets the material name for a part document.
        /// </summary>
        /// <param name="part">The part document.</param>
        /// <param name="configurationName">The configuration name.</param>
        /// <returns>The material name, or null if not assigned.</returns>
        public static string GetMaterial(PartDoc part, string configurationName)
        {
            if (part == null)
                return null;

            try
            {
                return part.GetMaterialPropertyName2(configurationName, out _);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Sets the material for a part document.
        /// </summary>
        /// <param name="part">The part document.</param>
        /// <param name="materialName">The name of the material to assign.</param>
        /// <param name="databasePath">The path to the material database.</param>
        /// <returns>True if the material was successfully assigned.</returns>
        public static bool SetMaterial(PartDoc part, string materialName, string databasePath)
        {
            if (part == null || string.IsNullOrEmpty(materialName))
                return false;

            try
            {
                part.SetMaterialPropertyName2("", databasePath, materialName);
                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Document Operations

        /// <summary>
        /// Gets the file path of a model document.
        /// </summary>
        /// <param name="model">The model document.</param>
        /// <returns>The full file path, or empty string if not saved.</returns>
        public static string GetFilePath(ModelDoc2 model)
        {
            return model?.GetPathName() ?? string.Empty;
        }

        /// <summary>
        /// Gets the document title (filename without path).
        /// </summary>
        /// <param name="model">The model document.</param>
        /// <returns>The document title.</returns>
        public static string GetTitle(ModelDoc2 model)
        {
            return model?.GetTitle() ?? string.Empty;
        }

        /// <summary>
        /// Checks if a document has been saved.
        /// </summary>
        /// <param name="model">The model document.</param>
        /// <returns>True if the document has been saved to disk.</returns>
        public static bool IsSaved(ModelDoc2 model)
        {
            return !string.IsNullOrEmpty(GetFilePath(model));
        }

        /// <summary>
        /// Saves a document.
        /// </summary>
        /// <param name="model">The model document to save.</param>
        /// <returns>True if the save operation was successful.</returns>
        public static bool Save(ModelDoc2 model)
        {
            if (model == null)
                return false;

            try
            {
                int errors = 0;
                int warnings = 0;
                return model.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref errors, ref warnings);
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Component Operations

        /// <summary>
        /// Resolves a lightweight component to fully loaded.
        /// </summary>
        /// <param name="component">The component to resolve.</param>
        /// <returns>True if the component was successfully resolved.</returns>
        public static bool ResolveComponent(Component2 component)
        {
            if (component == null)
                return false;

            try
            {
                component.SetLightweightToResolved();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Checks if a component is suppressed.
        /// </summary>
        /// <param name="component">The component to check.</param>
        /// <returns>True if the component is suppressed.</returns>
        public static bool IsComponentSuppressed(Component2 component)
        {
            return component?.IsSuppressed() ?? false;
        }

        /// <summary>
        /// Gets the model document from a component.
        /// </summary>
        /// <param name="component">The component.</param>
        /// <returns>The model document, or null if unavailable.</returns>
        public static ModelDoc2 GetModelFromComponent(Component2 component)
        {
            return component?.GetModelDoc2() as ModelDoc2;
        }

        #endregion
    }
}
