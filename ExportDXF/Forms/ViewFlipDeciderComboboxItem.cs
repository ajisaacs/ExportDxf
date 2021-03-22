using ExportDXF.ViewFlipDeciders;

namespace ExportDXF.Forms
{
    public class ViewFlipDeciderComboboxItem
    {
        public string Name { get; set; }

        public IViewFlipDecider ViewFlipDecider { get; set; }
    }
}