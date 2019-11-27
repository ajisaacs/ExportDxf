using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExportDXF
{
    public interface ItemExtractor
    {
        List<Item> GetItems();
    }

    public class BomItemExtractor : ItemExtractor
    {
        private BomTableAnnotation bom;
        private BomColumnIndices columnIndices;

        public BomItemExtractor(BomTableAnnotation bom)
        {
            this.bom = bom;
            columnIndices = new BomColumnIndices();
        }

        public bool SkipHiddenRows { get; set; }

        private void FindColumnIndices()
        {
            if (columnIndices == null)
                columnIndices = new BomColumnIndices();

            var table = bom as TableAnnotation;

            columnIndices.Description = table.IndexOfColumnTitle("Description");
            columnIndices.ItemNumber = table.IndexOfColumnType(swTableColumnTypes_e.swBomTableColumnType_ItemNumber);
            columnIndices.PartNumber = table.IndexOfColumnType(swTableColumnTypes_e.swBomTableColumnType_PartNumber);
            columnIndices.Quantity = table.IndexOfColumnType(swTableColumnTypes_e.swBomTableColumnType_Quantity);

            if (columnIndices.PartNumber == -1)
                throw new Exception("Part number column not found.");

            if (columnIndices.ItemNumber == -1)
                throw new Exception("Item number column not found.");

            if (columnIndices.Description == -1)
                throw new Exception("Description column not found.");

            if (columnIndices.Quantity == -1)
                throw new Exception("Quantity column not found.");
        }

        private Item GetItem(int rowIndex)
        {
            var item = new Item();
            var table = bom as TableAnnotation;

            if (columnIndices.ItemNumber != -1)
            {
                item.ItemNo = table.DisplayedText[rowIndex, columnIndices.ItemNumber];
            }

            if (columnIndices.PartNumber != -1)
            {
                item.PartName = table.DisplayedText[rowIndex, columnIndices.PartNumber];
            }

            if (columnIndices.Description != -1)
            {
                item.Description = table.DisplayedText[rowIndex, columnIndices.Description];
            }

            if (columnIndices.Quantity != -1)
            {
                var qtyString = table.DisplayedText[rowIndex, columnIndices.Quantity];

                int qty = 0;
                int.TryParse(qtyString, out qty);

                item.Quantity = qty;
            }

            item.Component = GetComponent(rowIndex);

            return item;
        }

        private Component2 GetComponent(int rowIndex)
        {
            var isBOMPartsOnly = bom.BomFeature.TableType == (int)swBomType_e.swBomType_PartsOnly;

            IEnumerable<Component2> components;

            if (isBOMPartsOnly)
            {
                components = ((Array)bom.GetComponents2(rowIndex, bom.BomFeature.Configuration))?.Cast<Component2>();
            }
            else
            {
                components = ((Array)bom.GetComponents(rowIndex))?.Cast<Component2>();
            }

            if (components == null)
                return null;

            foreach (var component in components)
            {
                component.SetLightweightToResolved();

                if (component.IsSuppressed())
                    continue;

                return component;
            }

            return null;
        }

        public List<Item> GetItems()
        {
            FindColumnIndices();

            var items = new List<Item>();
            var table = bom as TableAnnotation;

            for (int rowIndex = 1; rowIndex < table.RowCount; rowIndex++)
            {
                var isRowHidden = table.RowHidden[rowIndex];

                if (isRowHidden && SkipHiddenRows)
                    continue;

                var item = GetItem(rowIndex);
                items.Add(item);
            }

            return items;
        }
    }

    public class AssemblyItemExtractor : ItemExtractor
    {
        private AssemblyDoc assembly;

        public AssemblyItemExtractor(AssemblyDoc assembly)
        {
            this.assembly = assembly;
        }

        public bool TopLevelOnly { get; set; }

        private string GetComponentName(Component2 component)
        {
            var filepath = component.GetTitle();
            var filename = Path.GetFileNameWithoutExtension(filepath);
            var isDefaultConfig = component.ReferencedConfiguration.ToLower() == "default";

            return isDefaultConfig ? filename : $"{filename} [{component.ReferencedConfiguration}]";
        }

        public List<Item> GetItems()
        {
            var list = new List<Item>();

            assembly.ResolveAllLightWeightComponents(false);

            var assemblyComponents = ((Array)assembly.GetComponents(TopLevelOnly))
                .Cast<Component2>()
                .Where(c => !c.IsHidden(true));

            var componentGroups = assemblyComponents
                .GroupBy(c => c.GetTitle() + c.ReferencedConfiguration);

            foreach (var group in componentGroups)
            {
                var component = group.First();
                var model = component.GetModelDoc2() as ModelDoc2;

                if (model == null)
                    continue;

                var name = GetComponentName(component);

                list.Add(new Item
                {
                    PartName = name,
                    Quantity = group.Count(),
                    Component = component
                });
            }

            return list;
        }
    }

    public class BomColumnIndices
    {
        public int ItemNumber { get; set; } = -1;

        public int Quantity { get; set; } = -1;

        public int Description { get; set; } = -1;

        public int PartNumber { get; set; } = -1;
    }
}