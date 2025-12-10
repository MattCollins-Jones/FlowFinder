namespace Flow_Finder
{
    partial class FlowFinderControl
    {
        /// <summary> 
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary> 
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas 
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FlowFinderControl));
            this.toolStripMenu = new System.Windows.Forms.ToolStrip();
            this.tsbClose = new System.Windows.Forms.ToolStripButton();
            this.tssSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.tsbListCloudFlows = new System.Windows.Forms.ToolStripButton();
            this.tsbExport = new System.Windows.Forms.ToolStripButton();
            this.cmbSolutions = new System.Windows.Forms.ToolStripComboBox();
            this.dgvFlows = new System.Windows.Forms.DataGridView();
            this.toolStripMenu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvFlows)).BeginInit();
            this.SuspendLayout();
            // 
            // toolStripMenu
            // 
            this.toolStripMenu.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.toolStripMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsbClose,
            this.tssSeparator1,
            this.tsbListCloudFlows,
            this.tsbExport,
            this.cmbSolutions});
            this.toolStripMenu.Location = new System.Drawing.Point(0, 0);
            this.toolStripMenu.Name = "toolStripMenu";
            this.toolStripMenu.Size = new System.Drawing.Size(1095, 38);
            this.toolStripMenu.TabIndex = 4;
            this.toolStripMenu.Text = "toolStrip1";
            this.toolStripMenu.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.toolStripMenu_ItemClicked);
            // 
            // tsbClose
            // 
            this.tsbClose.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbClose.Name = "tsbClose";
            this.tsbClose.Size = new System.Drawing.Size(129, 33);
            this.tsbClose.Text = "Close this tool";
            this.tsbClose.Click += new System.EventHandler(this.tsbClose_Click);
            // 
            // tssSeparator1
            // 
            this.tssSeparator1.Name = "tssSeparator1";
            this.tssSeparator1.Size = new System.Drawing.Size(6, 38);
            // 
            // tsbListCloudFlows
            // 
            this.tsbListCloudFlows.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbListCloudFlows.Name = "tsbListCloudFlows";
            this.tsbListCloudFlows.Size = new System.Drawing.Size(144, 33);
            this.tsbListCloudFlows.Text = "List Cloud Flows";
            this.tsbListCloudFlows.Click += new System.EventHandler(this.tsbListCloudFlows_Click);
            // 
            // tsbExport
            // 
            this.tsbExport.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.tsbExport.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbExport.Margin = new System.Windows.Forms.Padding(10, 1, 10, 1);
            this.tsbExport.Name = "tsbExport";
            this.tsbExport.Size = new System.Drawing.Size(67, 36);
            this.tsbExport.Text = "Export";
            this.tsbExport.Click += new System.EventHandler(this.tsbExport_Click);
            // 
            // cmbSolutions
            // 
            this.cmbSolutions.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.cmbSolutions.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSolutions.Margin = new System.Windows.Forms.Padding(0, 1, 10, 1);
            this.cmbSolutions.Name = "cmbSolutions";
            this.cmbSolutions.Size = new System.Drawing.Size(310, 36);
            this.cmbSolutions.SelectedIndexChanged += new System.EventHandler(this.cmbSolutions_SelectedIndexChanged);
            // 
            // dgvFlows
            // 
            this.dgvFlows.AllowUserToAddRows = false;
            this.dgvFlows.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvFlows.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvFlows.Location = new System.Drawing.Point(12, 40);
            this.dgvFlows.MultiSelect = false;
            this.dgvFlows.Name = "dgvFlows";
            this.dgvFlows.ReadOnly = true;
            this.dgvFlows.RowHeadersWidth = 62;
            this.dgvFlows.RowTemplate.Height = 28;
            this.dgvFlows.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvFlows.Size = new System.Drawing.Size(1071, 406);
            this.dgvFlows.TabIndex = 5;
            // 
            // MyPluginControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.dgvFlows);
            this.Controls.Add(this.toolStripMenu);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "MyPluginControl";
            this.PluginIcon = ((System.Drawing.Icon)(resources.GetObject("$this.PluginIcon")));
            this.Size = new System.Drawing.Size(1095, 462);
            this.TabIcon = ((System.Drawing.Image)(resources.GetObject("$this.TabIcon")));
            this.ToolName = "Flow Finder";
            this.Load += new System.EventHandler(this.MyPluginControl_Load);
            this.toolStripMenu.ResumeLayout(false);
            this.toolStripMenu.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvFlows)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ToolStrip toolStripMenu;
        private System.Windows.Forms.ToolStripButton tsbClose;
        private System.Windows.Forms.ToolStripSeparator tssSeparator1;
        private System.Windows.Forms.ToolStripButton tsbListCloudFlows;
        private System.Windows.Forms.ToolStripButton tsbExport;
        private System.Windows.Forms.ToolStripComboBox cmbSolutions;
        private System.Windows.Forms.DataGridView dgvFlows;
    }
}
