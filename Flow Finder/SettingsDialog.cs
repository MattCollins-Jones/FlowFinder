using System;
using System.Windows.Forms;

namespace Flow_Finder
{
    internal sealed class SettingsDialog : Form
    {
        public Settings PluginSettings { get; private set; }
        private RadioButton rbNone; private RadioButton rbRefreshFlow; private RadioButton rbRefreshAll; private Button btnOk; private Button btnCancel;
        public SettingsDialog(Settings current) { PluginSettings = current ?? new Settings(); Initialize(); }
        private void Initialize()
        {
            this.Text = "Flow Finder Settings"; this.Width = 420; this.Height = 180; this.StartPosition = FormStartPosition.CenterParent;
            rbNone = new System.Windows.Forms.RadioButton() { Left = 10, Top = 10, Width = 380, Text = "Do not refresh after dialogs" };
            rbRefreshFlow = new System.Windows.Forms.RadioButton() { Left = 10, Top = 35, Width = 380, Text = "Refresh only the flow that was changed" };
            rbRefreshAll = new System.Windows.Forms.RadioButton() { Left = 10, Top = 60, Width = 380, Text = "Refresh the entire flows list" };
            this.Controls.Add(rbNone); this.Controls.Add(rbRefreshFlow); this.Controls.Add(rbRefreshAll);
            switch (PluginSettings?.RefreshAfterDialogMode ?? Flow_Finder.RefreshMode.RefreshFlow) { case Flow_Finder.RefreshMode.None: rbNone.Checked = true; break; case Flow_Finder.RefreshMode.RefreshFlow: rbRefreshFlow.Checked = true; break; case Flow_Finder.RefreshMode.RefreshAll: rbRefreshAll.Checked = true; break; default: rbRefreshFlow.Checked = true; break; }
            btnOk = new System.Windows.Forms.Button() { Left = 220, Top = 100, Width = 80, Text = "OK" };
            btnOk.Click += (s, e) => { if (rbNone.Checked) PluginSettings.RefreshAfterDialogMode = Flow_Finder.RefreshMode.None; else if (rbRefreshFlow.Checked) PluginSettings.RefreshAfterDialogMode = Flow_Finder.RefreshMode.RefreshFlow; else PluginSettings.RefreshAfterDialogMode = Flow_Finder.RefreshMode.RefreshAll; this.DialogResult = DialogResult.OK; this.Close(); };
            this.Controls.Add(btnOk);
            btnCancel = new System.Windows.Forms.Button() { Left = 310, Top = 100, Width = 80, Text = "Cancel" }; btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); }; this.Controls.Add(btnCancel);

            // Add a feedback button so users see feedback option in settings
            var btnFeedback = new System.Windows.Forms.Button() { Left = 10, Top = 100, Width = 200, Text = "Report feedback (GitHub)" };
            btnFeedback.Click += (s, e) => {
                try
                {
                    var asm = System.Reflection.Assembly.GetExecutingAssembly();
                    var ver = asm.GetName().Version?.ToString() ?? "unknown";
                    var title = System.Uri.EscapeDataString($"Issue: Flow Finder v{ver}");
                    var bodyBuilder = new System.Text.StringBuilder();
                    bodyBuilder.AppendLine("Please describe the issue you encountered and steps to reproduce:");
                    bodyBuilder.AppendLine();
                    bodyBuilder.AppendLine("---");
                    bodyBuilder.AppendLine($"Plugin version: {ver}");
                    bodyBuilder.AppendLine($"Assembly: {asm.GetName().Name}");
                    bodyBuilder.AppendLine($"OS: {Environment.OSVersion}");
                    bodyBuilder.AppendLine($"CLR: {Environment.Version}");
                    var body = System.Uri.EscapeDataString(bodyBuilder.ToString());
                    var url = $"https://github.com/{FlowFinderControl.GitHubUserName}/{FlowFinderControl.GitHubRepoName}/issues/new?title={title}&body={body}";
                    var psi = new System.Diagnostics.ProcessStartInfo { FileName = url, UseShellExecute = true };
                    System.Diagnostics.Process.Start(psi);
                }
                catch (Exception ex)
                {
                    try { MessageBox.Show("Failed to open the GitHub issues page: " + ex.Message); } catch { }
                }
            };
            this.Controls.Add(btnFeedback);
        }
    }
}
