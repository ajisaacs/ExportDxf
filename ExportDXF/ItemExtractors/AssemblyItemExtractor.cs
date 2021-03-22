using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExportDXF.ItemExtractors
{
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
                    Component = component,
                    Configuration = component.ReferencedConfiguration
                });
            }

            return list;
        }
    }
}