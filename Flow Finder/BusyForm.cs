using System.Windows.Forms;

namespace Flow_Finder
{
    internal sealed class BusyForm : Form
    {
        public BusyForm(string message = "Working...")
        {
            this.Width = 300; this.Height = 80; this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent; this.ControlBox = false;
            var lbl = new System.Windows.Forms.Label() { Left = 10, Top = 10, Width = 280, Text = message };
            var pb = new ProgressBar() { Left = 10, Top = 30, Width = 280, Style = ProgressBarStyle.Marquee, MarqueeAnimationSpeed = 30 };
            this.Controls.Add(lbl); this.Controls.Add(pb);
        }
    }
}
