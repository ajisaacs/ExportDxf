using System.Windows.Forms;

namespace ExportDXF
{
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
}
