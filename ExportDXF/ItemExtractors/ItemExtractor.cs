using System.Collections.Generic;

namespace ExportDXF.ItemExtractors
{
    public interface ItemExtractor
    {
        List<Item> GetItems();
    }
}