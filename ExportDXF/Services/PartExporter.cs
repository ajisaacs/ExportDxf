using ExportDXF.Extensions;
using ExportDXF.Models;
using ExportDXF.Utilities;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;

namespace ExportDXF.Services
{
    /// <summary>
    /// Service for exporting parts to DXF format.
    /// </summary>
    public interface IPartExporter
    {
        /// <summary>
        /// Exports a single part document to DXF.
        /// Returns an Item with export metadata (filename, hash, sheet metal properties), or null if export failed.
        /// </summary>
        /// <param name="part">The part document to export.</param>
        /// <param name="saveDirectory">The temp directory where the DXF file will be saved.</param>
        /// <param name="context">The export context.</param>
        Item ExportSinglePart(PartDoc part, string saveDirectory, ExportContext context);

        /// <summary>
        /// Exports an item (component from BOM or assembly) to DXF.
        /// </summary>
        /// <param name="item">The item to export.</param>
        /// <param name="saveDirectory">The temp directory where the DXF file will be saved.</param>
        /// <param name="context">The export context.</param>
        void ExportItem(Item item, string saveDirectory, ExportContext context);
    }

    public class PartExporter : IPartExporter
    {
        public PartExporter()
        {
        }

        public Item ExportSinglePart(PartDoc part, string saveDirectory, ExportContext context)
        {
            if (part == null)
                throw new ArgumentNullException(nameof(part));

            if (string.IsNullOrWhiteSpace(saveDirectory))
                throw new ArgumentException("Save directory cannot be null or empty.", nameof(saveDirectory));

            if (context == null)
                throw new ArgumentNullException(nameof(context));

            var model = part as ModelDoc2;
            var activeConfig = model.GetActiveConfiguration() as SolidWorks.Interop.sldworks.Configuration;
            var originalConfigName = activeConfig?.Name;

            try
            {
                var fileName = GetSinglePartFileName(model);
                var savePath = Path.Combine(saveDirectory, fileName + ".dxf");

                // Build result item with metadata
                var item = new Item
                {
                    PartName = model.GetTitle()?.Replace(".SLDPRT", "") ?? "",
                    Configuration = originalConfigName ?? "",
                    Quantity = 1
                };

                // Enrich with sheet metal properties and description
                var sheetMetalProps = SolidWorksHelper.GetSheetMetalProperties(model);
                if (sheetMetalProps != null)
                {
                    item.Thickness = sheetMetalProps.Thickness;
                    item.KFactor = sheetMetalProps.KFactor;
                    item.BendRadius = sheetMetalProps.BendRadius;
                }

                // Get description from custom properties
                var configPropMgr = model.Extension.CustomPropertyManager[originalConfigName];
                item.Description = configPropMgr?.Get("Description");
                if (string.IsNullOrEmpty(item.Description))
                {
                    var docPropMgr = model.Extension.CustomPropertyManager[""];
                    item.Description = docPropMgr?.Get("Description");
                }
                item.Description = TextHelper.RemoveXmlTags(item.Description);

                // Get material
                item.Material = part.GetMaterialPropertyName2(originalConfigName, out _);

                context.GetOrCreateTemplateDrawing();

                if (ExportPartToDxf(part, originalConfigName, savePath, context))
                {
                    item.FileName = Path.GetFileNameWithoutExtension(savePath);
                    item.ContentHash = Utilities.ContentHasher.ComputeDxfContentHash(savePath);
                    item.LocalTempPath = savePath;
                    return item;
                }
                else
                {
                    return null;
                }
            }
            finally
            {
                if (originalConfigName != null)
                {
                    model.ShowConfiguration(originalConfigName);
                }
            }
        }

        public void ExportItem(Item item, string saveDirectory, ExportContext context)
        {
            if (string.IsNullOrWhiteSpace(saveDirectory))
                throw new ArgumentException("Save directory cannot be null or empty.", nameof(saveDirectory));

            if (context == null)
                throw new ArgumentNullException(nameof(context));

            if (item?.Component == null)
            {
                context.ProgressCallback?.Invoke("Skipped, no component", LogLevel.Warning, $"Item {item?.ItemNo}");
                return;
            }

            context.CancellationToken.ThrowIfCancellationRequested();

            item.Component.SetLightweightToResolved();

            var model = item.Component.GetModelDoc2() as ModelDoc2;
            var part = model as PartDoc;

            if (part == null)
            {
                context.ProgressCallback?.Invoke("Skipped, not a part document", LogLevel.Info, item.PartName);
                return;
            }

            EnrichItemWithMetadata(item, model, part);

            var fileName = GetItemFileName(item);
            var savePath = Path.Combine(saveDirectory, fileName + ".dxf");

            var templateDrawing = context.GetOrCreateTemplateDrawing();

            if (ExportPartToDxf(part, item.Component.ReferencedConfiguration, savePath, context))
            {
                item.FileName = Path.GetFileNameWithoutExtension(savePath);
                item.ContentHash = Utilities.ContentHasher.ComputeDxfContentHash(savePath);
                item.LocalTempPath = savePath;
            }
            else
            {
                LogExportFailure(item, context);
            }
        }

        private void EnrichItemWithMetadata(Item item, ModelDoc2 model, PartDoc part)
        {
            // Get sheet metal properties
            var sheetMetalProps = SolidWorksHelper.GetSheetMetalProperties(model);
            if (sheetMetalProps != null)
            {
                item.Thickness = sheetMetalProps.Thickness;
                item.KFactor = sheetMetalProps.KFactor;
                item.BendRadius = sheetMetalProps.BendRadius;
            }

            // Get description from custom properties
            var config = item.Component.ReferencedConfiguration;

            // Try configuration-specific properties first
            var configPropertyManager = model.Extension.CustomPropertyManager[config];
            item.Description = configPropertyManager?.Get("Description");

            // Fall back to document-level properties if no config-specific description
            if (string.IsNullOrEmpty(item.Description))
            {
                var docPropertyManager = model.Extension.CustomPropertyManager[""];
                item.Description = docPropertyManager?.Get("Description");
            }

            item.Description = TextHelper.RemoveXmlTags(item.Description);

            // Get material
            item.Material = part.GetMaterialPropertyName2(config, out _);
        }

        private bool ExportPartToDxf(
            PartDoc part,
            string configName,
            string savePath,
            ExportContext context)
        {
            try
            {
                var model = part as ModelDoc2;

                var partTitle = model.GetTitle();

                if (!model.IsSheetMetal())
                {
                    context.ProgressCallback?.Invoke("Skipped, not sheet metal", LogLevel.Info, partTitle);
                    return false;
                }

                var templateDrawing = context.GetOrCreateTemplateDrawing();

                SolidWorksHelper.ConfigureFlatPatternSettings(model);

                var sheet = templateDrawing.IGetCurrentSheet();
                var modelName = Path.GetFileNameWithoutExtension(model.GetPathName());
                sheet.SetName(modelName);

                context.ProgressCallback?.Invoke("Creating flat pattern", LogLevel.Info, partTitle);

                var view = CreateFlatPatternView(templateDrawing, model, configName);

                if (view == null)
                {
                    context.ProgressCallback?.Invoke("Failed to create flat pattern", LogLevel.Error, partTitle);
                    return false;
                }

                ConfigureFlatPatternView(view, templateDrawing, model, configName, context);

                if (context.ViewFlipDecider?.ShouldFlip(view) == true)
                {
                    context.ProgressCallback?.Invoke("Flipped view", LogLevel.Info, partTitle);
                    view.FlipView = true;
                }

                var drawingModel = templateDrawing as ModelDoc2;
                drawingModel.SaveAs(savePath);

                AddEtchLines(savePath);

                context.ProgressCallback?.Invoke($"Saved to \"{savePath}\"", LogLevel.Info, partTitle);

                DeleteView(drawingModel, view);

                return true;
            }
            catch (Exception ex)
            {
                context.ProgressCallback?.Invoke($"Export failed: {ex.Message}", LogLevel.Error, null);
                return false;
            }
        }

        private SolidWorks.Interop.sldworks.View CreateFlatPatternView(
            DrawingDoc drawing,
            ModelDoc2 part,
            string configName)
        {
            return drawing.CreateFlatPatternViewFromModelView3(
                part.GetPathName(),
                configName,
                0, 0, 0,
                false,
                false);
        }

        private void ConfigureFlatPatternView(
            SolidWorks.Interop.sldworks.View view,
            DrawingDoc drawing,
            ModelDoc2 partModel,
            string configName,
            ExportContext context)
        {
            view.ShowSheetMetalBendNotes = true;

            var drawingModel = drawing as ModelDoc2;
            drawingModel.ViewZoomtofit2();

            var flatPatternModel = ViewHelper.GetModelFromView(view);
            SolidWorksHelper.SetFlatPatternSuppressionState(
                flatPatternModel,
                swComponentSuppressionState_e.swComponentFullyResolved);

            if (ViewHelper.HasSupressedBends(view))
            {
                var title = partModel.GetTitle();
                context.ProgressCallback?.Invoke("A bend is suppressed, please check flat pattern", LogLevel.Error, title);
            }

            if (ViewHelper.HideModelSketches(view))
            {
                // Recreate view without sketches
                DeleteView(drawingModel, view);

                view = CreateFlatPatternView(drawing, partModel, configName);
                view.ShowSheetMetalBendNotes = true;
            }
        }

        private void DeleteView(ModelDoc2 drawing, SolidWorks.Interop.sldworks.View view)
        {
            drawing.SelectByName(0, view.Name);
            drawing.DeleteSelection(false);
        }

        private void AddEtchLines(string dxfPath)
        {
            try
            {
                var etcher = new EtchBendLines.Etcher();
                etcher.AddEtchLines(dxfPath);
                FixDegreeSymbol(dxfPath);
            }
            catch (Exception)
            {
                // Silently fail if etch lines can't be added
            }
        }

        /// <summary>
        /// Workaround for ACadSharp encoding bug (no upstream fix as of v3.4.9).
        /// ACadSharp's DxfReader uses $DWGCODEPAGE (ANSI_1252) to decode text, but
        /// AC1018+ DXF files use UTF-8. The degree symbol ° (UTF-8: C2 B0) gets
        /// misread as two ANSI_1252 characters: Â (C2) and ° (B0).
        /// See: https://github.com/DomCR/ACadSharp/issues?q=encoding
        /// </summary>
        private static void FixDegreeSymbol(string path)
        {
            var text = System.IO.File.ReadAllText(path);
            if (text.Contains("\u00C2\u00B0"))
            {
                text = text.Replace("\u00C2\u00B0", "\u00B0");
                System.IO.File.WriteAllText(path, text);
            }
        }

        private string GetSinglePartFileName(ModelDoc2 model)
        {
            var title = model.GetTitle().Replace(".SLDPRT", "");
            var config = model.ConfigurationManager.ActiveConfiguration.Name;
            var isDefaultConfig = string.Equals(config, "default", StringComparison.OrdinalIgnoreCase);

            return isDefaultConfig ? title : $"{title} [{config}]";
        }

        private string GetItemFileName(Item item)
        {
            if (string.IsNullOrWhiteSpace(item.ItemNo))
                return item.PartName ?? "unknown";

            var num = item.ItemNo.PadLeft(2, '0');
            return $"PT{num}";
        }

        private void LogExportFailure(Item item, ExportContext context)
        {
            var desc = item.Description?.ToLower() ?? string.Empty;

            if (desc.Contains("laser"))
            {
                context.ProgressCallback?.Invoke(
                    "Export failed but description says it is laser cut",
                    LogLevel.Error,
                    item.PartName);
            }
            else if (desc.Contains("plasma"))
            {
                context.ProgressCallback?.Invoke(
                    "Export failed but description says it is plasma cut",
                    LogLevel.Error,
                    item.PartName);
            }
        }
    }
}
