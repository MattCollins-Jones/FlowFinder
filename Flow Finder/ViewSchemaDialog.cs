using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Flow_Finder
{
    internal class ViewSchemaDialog : Form
    {
        private readonly string _rawJson;
        private RichTextBox _rtbJson;
        private Button _btnCopy;
        private Button _btnClose;

        internal ViewSchemaDialog(string flowName, string rawJson)
        {
            _rawJson = rawJson;
            InitializeForm(flowName);
        }

        private void InitializeForm(string flowName)
        {
            Text = $"Flow Schema \u2014 {flowName}";
            Size = new Size(960, 700);
            MinimumSize = new Size(600, 400);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;

            _rtbJson = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 10f, FontStyle.Regular, GraphicsUnit.Point),
                ScrollBars = RichTextBoxScrollBars.Both,
                WordWrap = false,
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.None
            };

            var bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 44,
                Padding = new Padding(8, 6, 8, 6)
            };

            _btnCopy = new Button
            {
                Text = "Copy to Clipboard",
                AutoSize = true,
                Dock = DockStyle.Left
            };
            _btnCopy.Click += BtnCopy_Click;

            _btnClose = new Button
            {
                Text = "Close",
                Width = 80,
                Dock = DockStyle.Right,
                DialogResult = DialogResult.Cancel
            };

            bottomPanel.Controls.Add(_btnCopy);
            bottomPanel.Controls.Add(_btnClose);

            Controls.Add(_rtbJson);
            Controls.Add(bottomPanel);

            CancelButton = _btnClose;

            LoadJson();
        }

        private void LoadJson()
        {
            try
            {
                var pretty = JToken.Parse(_rawJson).ToString(Formatting.Indented);
                _rtbJson.Text = pretty;
            }
            catch
            {
                _rtbJson.Text = _rawJson;
            }
        }

        private void BtnCopy_Click(object sender, EventArgs e)
        {
            try
            {
                Clipboard.SetText(_rtbJson.Text);
                MessageBox.Show("Schema copied to clipboard.", "Copied", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to copy: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
