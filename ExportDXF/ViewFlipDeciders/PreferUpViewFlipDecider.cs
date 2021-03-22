using System.Linq;
using System.Windows.Forms;

namespace ExportDXF.ViewFlipDeciders
{
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