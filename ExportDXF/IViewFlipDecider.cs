namespace ExportDXF
{
	public interface IViewFlipDecider
    {
        bool ShouldFlip(SolidWorks.Interop.sldworks.View view);
    }
}
