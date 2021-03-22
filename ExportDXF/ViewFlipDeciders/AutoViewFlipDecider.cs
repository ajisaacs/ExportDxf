using System.Linq;

namespace ExportDXF.ViewFlipDeciders
{
    public class AutoViewFlipDecider : IViewFlipDecider
    {
        public string Name => "Automatic";

        public bool ShouldFlip(SolidWorks.Interop.sldworks.View view)
        {
            var orientation = ViewHelper.GetOrientation(view);
            var bounds = ViewHelper.GetBounds(view);
            var bends = ViewHelper.GetBends(view);

            var up = bends.Where(b => b.Direction == BendDirection.Up).ToList();
            var down = bends.Where(b => b.Direction == BendDirection.Down).ToList();

            if (down.Count == 0)
                return false;

            if (up.Count == 0)
                return true;

            var bend = ViewHelper.GetBendClosestToBounds(bounds, bends);

            return bend.Direction == BendDirection.Down;
        }
    }
}