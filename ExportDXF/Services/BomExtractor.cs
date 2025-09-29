using ExportDXF.Extensions;
using ExportDXF.ItemExtractors;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace ExportDXF.Services
{
    /// <summary>
    /// Service for extracting items from a Bill of Materials (BOM).
    /// </summary>
    public interface IBomExtractor
    {
        /// <summary>
        /// Extracts items from all BOM tables in a drawing document.
        /// </summary>
        /// <param name="drawing">The drawing document containing BOM tables.</param>
        /// <param name="progressCallback">Optional callback for progress updates.</param>
        /// <returns>A list of extracted items.</returns>
        List<Item> ExtractFromDrawing(DrawingDoc drawing, Action<string, Color?> progressCallback);
    }

    public class BomExtractor : IBomExtractor
    {
        public List<Item> ExtractFromDrawing(DrawingDoc drawing, Action<string, Color?> progressCallback)
        {
            if (drawing == null)
                throw new ArgumentNullException(nameof(drawing));

            var bomTables = drawing.GetBomTables();

            if (bomTables.Count == 0)
            {
                progressCallback?.Invoke("Error: Bill of materials not found.", Color.Red);
                return new List<Item>();
            }

            progressCallback?.Invoke($"Found {bomTables.Count} BOM table(s)", null);

            var allItems = new List<Item>();

            foreach (var bom in bomTables)
            {
                try
                {
                    var extractor = new BomItemExtractor(bom)
                    {
                        SkipHiddenRows = true
                    };

                    progressCallback?.Invoke($"Fetching components from {bom.BomFeature.Name}", null);

                    var items = extractor.GetItems();
                    allItems.AddRange(items);
                }
                catch (Exception ex)
                {
                    progressCallback?.Invoke($"Failed to extract from {bom.BomFeature.Name}: {ex.Message}", Color.Red);
                }
            }

            return allItems;
        }
    }
}