using System.Linq;
using System.Windows.Forms;

namespace ExportDXF
{
    public interface IViewFlipDecider
    {
        bool ShouldFlip(SolidWorks.Interop.sldworks.View view);

        string Name { get; }
    }

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

            var bend = ViewHelper.ClosestToBounds(bounds, bends);

            return bend.Direction == BendDirection.Down;
        }
    }

    public class AskViewFlipDecider : IViewFlipDecider
    {
        public string Name => "Ask to flip";

        public bool ShouldFlip(SolidWorks.Interop.sldworks.View view)
        {
            var bends = ViewHelper.GetBends(view);

            if (bends.Count == 0)
                return false;

            return MessageBox.Show("Flip view?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        }

        public override string ToString()
        {
            return Name;
        }
    }

    public class PreferUpViewFlipDecider : IViewFlipDecider
    {
        public string Name => "Prefer up bends, ask if up/down";

        public bool ShouldFlip(SolidWorks.Interop.sldworks.View view)
        {
            var bends = ViewHelper.GetBends(view);
            var up = bends.Where(b => b.Direction == BendDirection.Up).ToList();
            var down = bends.Where(b => b.Direction == BendDirection.Down).ToList();

            if (up.Count > 0 && down.Count > 0)
            {
                return MessageBox.Show("Flip view?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
            }
            else
            {
                return down.Count > 0;
            }
        }
    }
}
