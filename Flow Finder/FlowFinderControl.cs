using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Messages;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;
using System.Runtime.ExceptionServices;

namespace Flow_Finder
{
    public partial class FlowFinderControl : PluginControlBase, IGitHubPlugin
    {
        // Implement IGitHubPlugin so XrmToolBox will show the standard feedback -> New issue menu
        public string RepositoryName => "FlowFinder";
        public string UserName => "MattCollins-Jones";

        private Settings mySettings;
        private DataTable flowsTable;
        private List<FlowInfo> lastResults = new List<FlowInfo>();

        private class FlowInfo
        {
            public Guid Id { get; set; }
            public string Name { get; set; }
            public bool InSolution { get; set; }
            public string Solutions { get; set; }
            public string Owner { get; set; }
            public Guid OwnerId { get; set; }
            public string CoOwners { get; set; }
            public string Owners { get; set; }
            public List<Guid> PrincipalIds { get; set; }
            public string Description { get; set; }
            public string CreatedBy { get; set; }
            public string TriggerSource { get; set; }
            public string TriggerEntity { get; set; }
            public string OtherDataSources { get; set; }
        }

        public FlowFinderControl()
        {
            AppDomain.CurrentDomain.FirstChanceException += CurrentDomain_FirstChanceException;
            InitializeComponent();
            InitializeFlowsTable();
        }

        private void CurrentDomain_FirstChanceException(object sender, FirstChanceExceptionEventArgs e)
        {
            try
            {
                var log = Path.Combine(Path.GetTempPath(), "FlowFinder_FirstChance.log");
                File.AppendAllText(log, DateTime.Now.ToString("o") + " - " + e.Exception.ToString() + Environment.NewLine + Environment.NewLine);
            }
            catch { }
        }

        private void InitializeFlowsTable()
        {
            flowsTable = new DataTable();
            flowsTable.Columns.Add("Name");
            flowsTable.Columns.Add("Description");
            flowsTable.Columns.Add("Solutions");
            flowsTable.Columns.Add("Primary Owner");
            flowsTable.Columns.Add("Co-Owners");
            flowsTable.Columns.Add("Triggering Source");
            flowsTable.Columns.Add("Triggering Entity");
            flowsTable.Columns.Add("Other Data Sources");

            // Do not set DataSource or AutoSizeColumnsMode here — postpone to Load to avoid layout/resizing race conditions
        }

        private Guid GetGuidFromAttribute(Entity e, string attr)
        {
            if (e == null || string.IsNullOrEmpty(attr) || !e.Attributes.ContainsKey(attr)) return Guid.Empty;
            var val = e.Attributes[attr];
            if (val == null) return Guid.Empty;
            if (val is Guid) return (Guid)val;
            if (val is EntityReference) return ((EntityReference)val).Id;
            if (val is AliasedValue)
            {
                var av = (AliasedValue)val;
                if (av.Value is Guid) return (Guid)av.Value;
                if (av.Value is EntityReference) return ((EntityReference)av.Value).Id;
            }
            return Guid.Empty;
        }

        private string GetEntityReferenceName(Entity e, string attr)
        {
            if (e == null || string.IsNullOrEmpty(attr)) return null;
            try
            {
                var er = e.GetAttributeValue<EntityReference>(attr);
                return er == null ? null : er.Name;
            }
            catch { return null; }
        }

        private string GetStringSafe(Entity e, string attr)
        {
            if (e == null || string.IsNullOrEmpty(attr) || !e.Attributes.ContainsKey(attr)) return null;
            var val = e.Attributes[attr];
            return val == null ? null : val.ToString();
        }

        private void ParseClientData(string clientJson, FlowInfo fi)
        {
            if (string.IsNullOrWhiteSpace(clientJson) || fi == null) return;
            try
            {
                var doc = Newtonsoft.Json.Linq.JObject.Parse(clientJson);
                var def = doc.SelectToken("properties.definition");
                if (def == null) return;

                var triggers = def["triggers"] as Newtonsoft.Json.Linq.JObject;
                if (triggers != null)
                {
                    var first = triggers.Properties().FirstOrDefault();
                    if (first != null)
                    {
                        var trigObj = first.Value as Newtonsoft.Json.Linq.JObject;
                        var type = trigObj?["type"]?.ToString();
                        fi.TriggerSource = MapTriggerTypeToSource(type, trigObj);

                        var inputs = trigObj?["inputs"] as Newtonsoft.Json.Linq.JObject;
                        if (inputs != null)
                        {
                            var host = inputs["host"] as Newtonsoft.Json.Linq.JObject;
                            if (host != null)
                            {
                                var apiId = host["apiId"]?.ToString();
                                if (!string.IsNullOrEmpty(apiId)) fi.TriggerSource = MapApiIdToSource(apiId);

                                var parameters = inputs["parameters"] as Newtonsoft.Json.Linq.JObject;
                                if (parameters != null)
                                {
                                    var entity = parameters.Properties().FirstOrDefault(p => p.Name.ToLower().Contains("entity") || p.Name.ToLower().Contains("entityname"));
                                    if (entity != null) fi.TriggerEntity = entity.Value.ToString();
                                }
                            }
                        }
                    }
                }

                var connRefs = doc.SelectToken("properties.connectionReferences") as Newtonsoft.Json.Linq.JObject;
                if (connRefs != null)
                {
                    var names = connRefs.Properties().Select(p => p.Name).Distinct().ToList();
                    fi.OtherDataSources = string.Join(", ", names);
                }
            }
            catch { }
        }

        private string MapTriggerTypeToSource(string type, Newtonsoft.Json.Linq.JObject trigObj)
        {
            if (string.IsNullOrEmpty(type)) return null;
            type = type.ToLowerInvariant();
            if (type.Contains("openapiconnectionwebhook") || type.Contains("openapiconnection")) return "Connector";
            if (type.Contains("request")) return "Manual/HTTP";
            if (type.Contains("recurrence")) return "Recurrence";
            if (type.Contains("workflow")) return "Workflow";
            return type;
        }

        private string MapApiIdToSource(string apiId)
        {
            if (string.IsNullOrEmpty(apiId)) return null;
            apiId = apiId.ToLowerInvariant();
            if (apiId.Contains("/providers/microsoft.powerapps/apis/shared_commondataserviceforapps")) return "Dataverse";
            if (apiId.Contains("/providers/microsoft.powerapps/apis/shared_sharepointonline")) return "SharePoint";
            if (apiId.Contains("/providers/microsoft.powerapps/apis/sharedoffice_365")) return "Office365";
            if (apiId.Contains("/providers/microsoft.powerapps/apis/http")) return "HTTP";
            return apiId.Split('/').LastOrDefault();
        }

        private string EscapeCsv(string s)
        {
            if (s == null) return string.Empty;
            if (s.Contains(",") || s.Contains('"') || s.Contains('\n')) return '"' + s.Replace("\"", "\"\"") + '"';
            return s;
        }

        private void ApplySolutionFilter(string filterText)
        {
            if (string.IsNullOrWhiteSpace(filterText) || filterText == "All solutions")
            {
                // Clear any existing row filter so the DefaultView shows all rows again
                try { if (flowsTable != null) flowsTable.DefaultView.RowFilter = string.Empty; } catch { }
                dgvFlows.DataSource = flowsTable;
                return;
            }

            var dv = flowsTable.DefaultView;
            dv.RowFilter = $"Solutions LIKE '%{filterText.Replace("'","''")}%'";
        }

        private void cmbSolutions_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSolutions.SelectedIndex < 0) return;
            var sel = cmbSolutions.ComboBox.Items[cmbSolutions.SelectedIndex];
            if (sel == null) return;
            var solutionName = sel.ToString();
            ApplySolutionFilter(solutionName);
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            FindCloudFlows();
        }

        private void FindCloudFlows()
        {
            // Ensure we're connected
            try
            {
                if (Service == null || ConnectionDetail == null || string.IsNullOrEmpty(ConnectionDetail.WebApplicationUrl))
                {
                    // Not connected — launch XrmToolBox connection window so user can connect
                    TryOpenConnectionDialog();
                    return;
                }
            }
            catch { /* fallback to safety */ }
             WorkAsync(new WorkAsyncInfo
             {
                Message = "Retrieving cloud flows and solutions",
                Work = (worker, args) =>
                {
                    LogInfo("Starting retrieval of cloud flows and solutions");
                    var qe = new QueryExpression("workflow") { ColumnSet = new ColumnSet("workflowid", "name", "ownerid", "type", "category", "description", "createdby"), Criteria = new FilterExpression() };
                    qe.Criteria.AddCondition("type", ConditionOperator.Equal, 1);
                    qe.Criteria.AddCondition("category", ConditionOperator.Equal, 6);
                    var flows = Service.RetrieveMultiple(qe);

                    if (flows == null || flows.Entities.Count == 0)
                    {
                        var qe2 = new QueryExpression("workflow") { ColumnSet = new ColumnSet("workflowid", "name", "ownerid", "type", "category", "description", "createdby"), Criteria = new FilterExpression() };
                        qe2.Criteria.AddCondition("type", ConditionOperator.Equal, 1);
                        qe2.Criteria.AddCondition("category", ConditionOperator.In, new object[] { 5, 6, 7 });
                        flows = Service.RetrieveMultiple(qe2);
                    }
                    if (flows == null || flows.Entities.Count == 0)
                    {
                        var qe3 = new QueryExpression("workflow") { ColumnSet = new ColumnSet("workflowid", "name", "ownerid", "type", "category", "description", "createdby") };
                        qe3.Criteria.AddCondition("type", ConditionOperator.Equal, 1);
                        flows = Service.RetrieveMultiple(qe3);
                    }

                    var scQ = new QueryExpression("solutioncomponent") { ColumnSet = new ColumnSet("solutionid", "objectid", "componenttype"), Criteria = new FilterExpression() };
                    scQ.Criteria.AddCondition("componenttype", ConditionOperator.Equal, 29);
                    var solutionComponents = Service.RetrieveMultiple(scQ);

                    var solutionIds = solutionComponents.Entities.Select(sc => GetGuidFromAttribute(sc, "solutionid")).Where(g => g != Guid.Empty).Distinct().ToList();
                    var solutionNames = new Dictionary<Guid, string>();
                    if (solutionIds.Any())
                    {
                        var solFetch = new QueryExpression("solution") { ColumnSet = new ColumnSet("solutionid", "friendlyname", "uniquename") };
                        solFetch.Criteria.AddCondition("solutionid", ConditionOperator.In, solutionIds.Select(g => (object)g).ToArray());
                        var sols = Service.RetrieveMultiple(solFetch);
                        foreach (var s in sols.Entities)
                        {
                            var friendly = GetStringSafe(s, "friendlyname") ?? GetStringSafe(s, "uniquename");
                            var uniq = GetStringSafe(s, "uniquename") ?? string.Empty;
                            if (!string.IsNullOrEmpty(uniq) && uniq.Equals("default", StringComparison.OrdinalIgnoreCase)) continue;
                            if (!string.IsNullOrEmpty(friendly) && (friendly.IndexOf("default solution", StringComparison.OrdinalIgnoreCase) >= 0 || friendly.IndexOf("active solution", StringComparison.OrdinalIgnoreCase) >= 0)) continue;
                            solutionNames[s.Id] = friendly;
                        }
                    }

                    var flowSolutions = new Dictionary<Guid, List<string>>();
                    foreach (var sc in solutionComponents.Entities)
                    {
                        var objId = GetGuidFromAttribute(sc, "objectid");
                        var solId = GetGuidFromAttribute(sc, "solutionid");
                        if (objId == Guid.Empty || solId == Guid.Empty) continue;
                        if (!flowSolutions.ContainsKey(objId)) flowSolutions[objId] = new List<string>();
                        if (solutionNames.ContainsKey(solId)) flowSolutions[objId].Add(solutionNames[solId]);
                    }

                    var allSolutionNames = flowSolutions.SelectMany(kvp => kvp.Value).Distinct().OrderBy(s => s).ToList();

                    var results = new List<FlowInfo>();
                    var principalIds = new HashSet<Guid>();

                    foreach (var f in flows.Entities)
                    {
                        var fi = new FlowInfo
                        {
                            Id = f.Id,
                            Name = GetStringSafe(f, "name"),
                            Owner = GetEntityReferenceName(f, "ownerid"),
                            OwnerId = GetGuidFromAttribute(f, "ownerid"),
                            Description = GetStringSafe(f, "description"),
                            CreatedBy = GetEntityReferenceName(f, "createdby")
                        };

                        if (flowSolutions.ContainsKey(f.Id)) { fi.InSolution = true; fi.Solutions = string.Join(", ", flowSolutions[f.Id].Distinct()); }
                        else fi.InSolution = false;

                        try
                        {
                            var req = new RetrieveSharedPrincipalsAndAccessRequest { Target = new EntityReference("workflow", f.Id) };
                            var resp = (RetrieveSharedPrincipalsAndAccessResponse)Service.Execute(req);
                            var principalGuids = new List<Guid>();
                            if (resp?.PrincipalAccesses != null)
                            {
                                foreach (var pa in resp.PrincipalAccesses)
                                {
                                    var principalRef = pa.Principal as EntityReference;
                                    if (principalRef != null)
                                    {
                                        principalGuids.Add(principalRef.Id);
                                        principalIds.Add(principalRef.Id);
                                    }
                                }
                            }
                            fi.PrincipalIds = principalGuids;
                        }
                        catch { }

                        results.Add(fi);
                    }

                    var principalNames = new Dictionary<Guid, string>();
                    if (principalIds.Any())
                    {
                        var idsArray = principalIds.ToArray();
                        var userQ = new QueryExpression("systemuser") { ColumnSet = new ColumnSet("systemuserid", "fullname") };
                        userQ.Criteria.AddCondition("systemuserid", ConditionOperator.In, idsArray.Cast<object>().ToArray());
                        var users = Service.RetrieveMultiple(userQ);
                        foreach (var u in users.Entities) principalNames[u.Id] = GetStringSafe(u, "fullname");

                        var teamQ = new QueryExpression("team") { ColumnSet = new ColumnSet("teamid", "name") };
                        teamQ.Criteria.AddCondition("teamid", ConditionOperator.In, idsArray.Cast<object>().ToArray());
                        var teams = Service.RetrieveMultiple(teamQ);
                        foreach (var t in teams.Entities) principalNames[t.Id] = GetStringSafe(t, "name");
                    }

                    foreach (var fi in results)
                    {
                        if (fi.PrincipalIds != null && fi.PrincipalIds.Any())
                        {
                            var coIds = fi.PrincipalIds.Where(id => id != fi.OwnerId).Distinct().ToList();
                            var coNames = coIds.Select(id => principalNames.ContainsKey(id) ? principalNames[id] : id.ToString()).ToList();
                            fi.CoOwners = coNames.Any() ? string.Join(", ", coNames) : null;
                        }
                        else fi.CoOwners = null;
                    }

                    foreach (var fi in results)
                    {
                        try
                        {
                            var wf = Service.Retrieve("workflow", fi.Id, new ColumnSet("clientdata"));
                            if (wf != null && wf.Attributes.ContainsKey("clientdata"))
                            {
                                var clientJson = GetStringSafe(wf, "clientdata");
                                ParseClientData(clientJson, fi);
                            }
                        }
                        catch { }
                    }

                    args.Result = new { Results = results, Solutions = allSolutionNames };
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        MessageBox.Show(args.Error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    dynamic r = args.Result;
                    lastResults = ((List<FlowInfo>)r.Results).OrderBy(f => f.Name).ToList();
                    var solutionNames = (List<string>)r.Solutions;

                    try { dgvFlows.DataSource = null; } catch { }
                    flowsTable.Clear();
                    foreach (var f in lastResults)
                    {
                        var row = flowsTable.NewRow();
                        row["Name"] = f.Name ?? string.Empty;
                        row["Description"] = f.Description ?? string.Empty;
                        row["Solutions"] = f.Solutions ?? string.Empty;
                        row["Primary Owner"] = f.Owner ?? string.Empty;
                        row["Co-Owners"] = f.CoOwners ?? string.Empty;
                        row["Triggering Source"] = f.TriggerSource ?? string.Empty;
                        row["Triggering Entity"] = f.TriggerEntity ?? string.Empty;
                        row["Other Data Sources"] = f.OtherDataSources ?? string.Empty;
                        flowsTable.Rows.Add(row);
                    }
                    dgvFlows.DataSource = flowsTable;

                    for (int i = 0; i < dgvFlows.Rows.Count && i < lastResults.Count; i++)
                    {
                        try { dgvFlows.Rows[i].Tag = lastResults[i].Id; } catch { }
                    }

                    cmbSolutions.SelectedIndexChanged -= cmbSolutions_SelectedIndexChanged;
                    cmbSolutions.SelectedIndexChanged -= cmbSolutions_SelectedIndexChanged;
                    try { cmbSolutions.ComboBox.DataSource = null; } catch { }
                    cmbSolutions.ComboBox.Items.Clear(); cmbSolutions.ComboBox.Items.Add("All solutions");
                    foreach (var sname in solutionNames) cmbSolutions.ComboBox.Items.Add(sname);
                    try { cmbSolutions.SelectedIndex = 0; } catch { cmbSolutions.ComboBox.SelectedIndex = 0; }
                    cmbSolutions.SelectedIndexChanged += cmbSolutions_SelectedIndexChanged;
                }
            });
        }

        // Designer event handlers
        private void MyPluginControl_Load(object sender, EventArgs e)
        {
            // minimal load: load settings if present
            if (!SettingsManager.Instance.TryLoad(GetType(), out mySettings))
            {
                mySettings = new Settings();
                LogWarning("Settings not found => a new settings file has been created!");
            }
            else
            {
                LogInfo("Settings found and loaded");
            }

            // Ensure runtime toolstrip buttons for manage and settings
            try
            {
                if (toolStripMenu != null)
                {
                    AddRuntimeButtonIfMissing("tsbManageCoOwners", "Manage Co-Owners", tsbManageCoOwners_Click);
                    AddRuntimeButtonIfMissing("tsbManageSolutions", "Manage Solutions", tsbManageSolutions_Click);
                    AddRuntimeButtonIfMissing("tsbSettingsRuntime", "Settings", tsbSettings_Click);
                    // Feedback is provided via the host Feedback menu (IGitHubPlugin). Do not add a runtime toolbar button here.
                 }
             }
             catch { }

            // Ensure filter shows immediately
            try
            {
                if (cmbSolutions != null)
                {
                    try { cmbSolutions.ComboBox.Items.Clear(); } catch { }
                    cmbSolutions.ComboBox.Items.Add("All solutions");
                    cmbSolutions.SelectedIndex = 0;
                }
            }
            catch { }

            // Assign DataSource and column sizing after control is created to avoid DataGridView auto-fill resize errors
            try
            {
                this.BeginInvoke((Action)(() =>
                {
                    try
                    {
                        if (dgvFlows != null)
                        {
                            dgvFlows.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                            if (flowsTable != null) dgvFlows.DataSource = flowsTable;
                        }
                    }
                    catch { }
                }));
            }
            catch { }
        }

        private void AddRuntimeButtonIfMissing(string name, string text, EventHandler handler)
        {
            bool has = false;
            foreach (ToolStripItem it in toolStripMenu.Items)
            {
                if (string.Equals(it.Name, name, StringComparison.OrdinalIgnoreCase)) { has = true; break; }
            }
            if (!has)
            {
                var btn = new ToolStripButton() { DisplayStyle = ToolStripItemDisplayStyle.Text, Name = name, Text = text };
                btn.Click += handler;
                toolStripMenu.Items.Add(new ToolStripSeparator());
                toolStripMenu.Items.Add(btn);
            }
        }

        private void tsbClose_Click(object sender, EventArgs e)
        {
            CloseTool();
        }

        private void tsbListCloudFlows_Click(object sender, EventArgs e)
        {
            // Use ExecuteMethod so the host can handle connection prompting (same pattern as the working plugin)
            ExecuteMethod(FindCloudFlows);
        }

        private void tsbExport_Click(object sender, EventArgs e)
        {
            if (dgvFlows == null || dgvFlows.Rows.Count == 0) { MessageBox.Show("No data to export."); return; }

            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                sfd.FileName = "FlowsExport.csv";
                if (sfd.ShowDialog() != DialogResult.OK) return;

                using (var writer = new StreamWriter(sfd.FileName, false, Encoding.UTF8))
                {
                    writer.WriteLine("Name,Description,Solutions,Primary Owner,Co-Owners,Triggering Source,Triggering Entity,Other Data Sources");
                    foreach (DataGridViewRow row in dgvFlows.Rows)
                    {
                        if (row.IsNewRow) continue;
                        var vals = new List<string>();
                        for (int i = 0; i < 8; i++) vals.Add(EscapeCsv(Convert.ToString(row.Cells[i].Value ?? string.Empty)));
                        writer.WriteLine(string.Join(",", vals));
                    }
                }

                MessageBox.Show("Export complete.");
            }
        }

        private void tsbSettings_Click(object sender, EventArgs e)
        {
            using (var dlg = new SettingsDialog(mySettings))
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    mySettings = dlg.PluginSettings;
                    SettingsManager.Instance.Save(GetType(), mySettings);
                    LogInfo("Settings updated and saved");
                }
            }
        }

        private class SettingsDialog : Form
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
                        var url = $"https://github.com/{"MattCollins-Jones"}/{"FlowFinder"}/issues/new?title={title}&body={body}";
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

        private void tsbManageCoOwners_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = null;
            if (dgvFlows.SelectedRows != null && dgvFlows.SelectedRows.Count > 0) row = dgvFlows.SelectedRows[0];
            else row = dgvFlows.CurrentRow;
            if (row == null) { MessageBox.Show("Please select a flow in the list first.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (row.Tag == null || !(row.Tag is Guid flowId)) return;
            using (var dlg = new ManageCoOwnersDialog(Service, flowId, s => LogInfo(s), s => LogWarning(s))) { dlg.ShowDialog(); }
            // refresh according to settings
            RefreshAccordingToSetting((Guid)row.Tag);
        }

        private void tsbManageSolutions_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = null;
            if (dgvFlows.SelectedRows != null && dgvFlows.SelectedRows.Count > 0) row = dgvFlows.SelectedRows[0];
            else row = dgvFlows.CurrentRow;
            if (row == null) { MessageBox.Show("Please select a flow in the list first.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (row.Tag == null || !(row.Tag is Guid flowId)) return;
            using (var dlg = new ManageSolutionsDialog(Service, flowId, s => LogInfo(s), s => LogWarning(s))) { dlg.ShowDialog(); }
            // refresh according to settings
            RefreshAccordingToSetting((Guid)row.Tag);
        }

        private void RefreshAccordingToSetting(Guid flowId)
        {
            var mode = mySettings?.RefreshAfterDialogMode ?? Flow_Finder.RefreshMode.RefreshFlow;
            if (mode == Flow_Finder.RefreshMode.RefreshAll) ExecuteMethod(FindCloudFlows);
            else if (mode == Flow_Finder.RefreshMode.RefreshFlow) RefreshSingleFlow(flowId);
        }

        private void RefreshSingleFlow(Guid flowId)
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Refreshing flow",
                Work = (w, a) =>
                {
                    var fi = new FlowInfo();
                    try
                    {
                        var wf = Service.Retrieve("workflow", flowId, new ColumnSet("workflowid", "name", "ownerid", "description", "createdby", "clientdata"));
                        fi.Id = wf.Id; fi.Name = GetStringSafe(wf, "name"); fi.Description = GetStringSafe(wf, "description"); fi.Owner = GetEntityReferenceName(wf, "ownerid"); fi.OwnerId = GetGuidFromAttribute(wf, "ownerid"); fi.CreatedBy = GetEntityReferenceName(wf, "createdby");

                        try
                        {
                            var req = new RetrieveSharedPrincipalsAndAccessRequest { Target = new EntityReference("workflow", flowId) };
                            var resp = (RetrieveSharedPrincipalsAndAccessResponse)Service.Execute(req);
                            var principalIds = new List<Guid>();
                            if (resp?.PrincipalAccesses != null) foreach (var pa in resp.PrincipalAccesses) if (pa.Principal is EntityReference er) principalIds.Add(er.Id);
                            if (principalIds.Any())
                            {
                                var userQ = new QueryExpression("systemuser") { ColumnSet = new ColumnSet("systemuserid", "fullname") };
                                userQ.Criteria.AddCondition("systemuserid", ConditionOperator.In, principalIds.Cast<object>().ToArray());
                                var users = Service.RetrieveMultiple(userQ);
                                var names = new Dictionary<Guid, string>();
                                foreach (var u in users.Entities) names[u.Id] = u.GetAttributeValue<string>("fullname");
                                var coIds = principalIds.Where(id => id != fi.OwnerId).Distinct().ToList();
                                var coNames = coIds.Select(id => names.ContainsKey(id) ? names[id] : id.ToString()).ToList();
                                fi.CoOwners = coNames.Any() ? string.Join(", ", coNames) : null;
                            }
                        }
                        catch { }

                        try
                        {
                            var scQ = new QueryExpression("solutioncomponent") { ColumnSet = new ColumnSet("solutionid", "objectid", "componenttype") };
                            scQ.Criteria.AddCondition("componenttype", ConditionOperator.Equal, 29);
                            scQ.Criteria.AddCondition("objectid", ConditionOperator.Equal, flowId);
                            var scRes = Service.RetrieveMultiple(scQ);
                            var solIds = scRes.Entities.Select(sc => GetGuidFromAttribute(sc, "solutionid")).Where(g => g != Guid.Empty).Distinct().ToArray();
                            if (solIds.Any())
                            {
                                var solQ = new QueryExpression("solution") { ColumnSet = new ColumnSet("solutionid", "friendlyname", "uniquename") };
                                solQ.Criteria.AddCondition("solutionid", ConditionOperator.In, solIds.Cast<object>().ToArray());
                                var sols = Service.RetrieveMultiple(solQ);
                                var names = sols.Entities.Select(s => s.GetAttributeValue<string>("friendlyname") ?? s.GetAttributeValue<string>("uniquename")).Where(n => !string.IsNullOrEmpty(n) && n.IndexOf("active solution", StringComparison.OrdinalIgnoreCase) < 0 && n.IndexOf("default solution", StringComparison.OrdinalIgnoreCase) < 0).ToList();
                                fi.Solutions = names.Any() ? string.Join(", ", names) : null;
                            }
                        }
                        catch { }

                        try
                        {
                            if (wf != null && wf.Attributes.ContainsKey("clientdata")) ParseClientData(GetStringSafe(wf, "clientdata"), fi);
                        }
                        catch { }
                    }
                    catch (Exception ex) { a.Result = ex; return; }
                    a.Result = fi;
                },
                PostWorkCallBack = (a) =>
                {
                    if (a.Error != null) { MessageBox.Show(a.Error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                    if (a.Result is Exception ex) { MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                    var fi = (FlowInfo)a.Result; var idx = lastResults.FindIndex(x => x.Id == fi.Id); if (idx >= 0) lastResults[idx] = fi; foreach (DataGridViewRow row in dgvFlows.Rows) { if (row.Tag is Guid id && id == fi.Id) { row.Cells[0].Value = fi.Name; row.Cells[1].Value = fi.Description ?? ""; row.Cells[2].Value = fi.Solutions ?? ""; row.Cells[3].Value = fi.Owner ?? ""; row.Cells[4].Value = fi.CoOwners ?? ""; try { row.Cells[4].ToolTipText = fi.CoOwners ?? string.Empty; } catch { } break; } }
                }
            });
        }

        private class BusyForm : Form { public BusyForm(string message = "Working...") { this.Width = 300; this.Height = 80; this.FormBorderStyle = FormBorderStyle.FixedDialog; this.StartPosition = FormStartPosition.CenterParent; this.ControlBox = false; var lbl = new System.Windows.Forms.Label() { Left = 10, Top = 10, Width = 280, Text = message }; var pb = new ProgressBar() { Left = 10, Top = 30, Width = 280, Style = ProgressBarStyle.Marquee, MarqueeAnimationSpeed = 30 }; this.Controls.Add(lbl); this.Controls.Add(pb); } }

        private class ManageCoOwnersDialog : Form
        {
            private IOrganizationService _service; private Guid _flowId; private Action<string> _logInfo; private Action<string> _logWarn;
            private ListBox lbCoOwners; private ComboBox cbUsers; private Button btnAdd; private Button btnRemove; private Button btnClose;

            public ManageCoOwnersDialog(IOrganizationService service, Guid flowId, Action<string> logInfo = null, Action<string> logWarn = null)
            {
                _service = service; _flowId = flowId; _logInfo = logInfo ?? (_ => { }); _logWarn = logWarn ?? (_ => { }); Initialize(); LoadData();
            }

            private void Initialize()
            {
                this.Text = "Manage Co-owners"; this.Width = 600; this.Height = 400; this.StartPosition = FormStartPosition.CenterParent;
                lbCoOwners = new ListBox() { Left = 10, Top = 10, Width = 350, Height = 300 }; this.Controls.Add(lbCoOwners);
                cbUsers = new ComboBox() { Left = 370, Top = 10, Width = 200, DropDownStyle = ComboBoxStyle.DropDown, AutoCompleteMode = AutoCompleteMode.SuggestAppend, AutoCompleteSource = AutoCompleteSource.ListItems }; this.Controls.Add(cbUsers);
                btnAdd = new Button() { Left = 370, Top = 50, Width = 200, Text = "Add as Co-owner" }; btnAdd.Click += BtnAdd_Click; this.Controls.Add(btnAdd);
                btnRemove = new Button() { Left = 370, Top = 90, Width = 200, Text = "Remove Selected Co-owner" }; btnRemove.Click += BtnRemove_Click; this.Controls.Add(btnRemove);
                btnClose = new Button() { Left = 370, Top = 300, Width = 200, Text = "Close" }; btnClose.Click += (s, e) => this.Close(); this.Controls.Add(btnClose);
            }

            private void LoadData()
            {
                if (this.InvokeRequired) { this.BeginInvoke((Action)LoadData); return; }
                lbCoOwners.Items.Clear(); cbUsers.DataSource = null; cbUsers.Items.Clear();
                var principalIds = new List<Guid>();
                try
                {
                    var req = new RetrieveSharedPrincipalsAndAccessRequest { Target = new EntityReference("workflow", _flowId) };
                    var resp = (RetrieveSharedPrincipalsAndAccessResponse)_service.Execute(req);
                    if (resp?.PrincipalAccesses != null) foreach (var pa in resp.PrincipalAccesses) if (pa.Principal is EntityReference er) principalIds.Add(er.Id);
                }
                catch (Exception ex) { _logWarn("Failed to load co-owners: " + ex.Message); MessageBox.Show("Failed to load co-owners: " + ex.Message); }

                var names = new Dictionary<Guid, string>();
                if (principalIds.Any())
                {
                    try { var userQ = new QueryExpression("systemuser") { ColumnSet = new ColumnSet("systemuserid", "fullname") }; userQ.Criteria.AddCondition("systemuserid", ConditionOperator.In, principalIds.Cast<object>().ToArray()); var users = _service.RetrieveMultiple(userQ); foreach (var u in users.Entities) names[u.Id] = u.GetAttributeValue<string>("fullname"); }
                    catch (Exception ex) { _logWarn("Failed to resolve co-owner user names: " + ex.Message); }
                    try { var teamQ = new QueryExpression("team") { ColumnSet = new ColumnSet("teamid", "name") }; teamQ.Criteria.AddCondition("teamid", ConditionOperator.In, principalIds.Cast<object>().ToArray()); var teams = _service.RetrieveMultiple(teamQ); foreach (var t in teams.Entities) names[t.Id] = t.GetAttributeValue<string>("name"); }
                    catch (Exception ex) { _logWarn("Failed to resolve team names: " + ex.Message); }
                }

                foreach (var id in principalIds) lbCoOwners.Items.Add(new ListItem { Id = id, Name = names.ContainsKey(id) ? names[id] : id.ToString() });

                // load users for candidates
                try
                {
                    Guid ownerId = Guid.Empty;
                    try
                    {
                        var wf = _service.Retrieve("workflow", _flowId, new ColumnSet("ownerid"));
                        if (wf != null && wf.Contains("ownerid"))
                        {
                            var or = wf.GetAttributeValue<EntityReference>("ownerid");
                            if (or != null) ownerId = or.Id;
                        }
                    }
                    catch { }

                    var exclude = new HashSet<Guid>(principalIds);
                    if (ownerId != Guid.Empty) exclude.Add(ownerId);

                    var userList = new List<ListItem>();
                    var q = new QueryExpression("systemuser") { ColumnSet = new ColumnSet("systemuserid", "fullname") };
                    var res = _service.RetrieveMultiple(q);
                    foreach (var u in res.Entities)
                    {
                        var id = u.Id;
                        if (exclude.Contains(id)) continue;
                        userList.Add(new ListItem { Id = id, Name = u.GetAttributeValue<string>("fullname") });
                    }

                    userList = userList.OrderBy(x => x.Name).ToList();
                    cbUsers.DisplayMember = "Name"; cbUsers.ValueMember = "Id";
                    cbUsers.DataSource = userList;
                    _logInfo($"Loaded {userList.Count} candidate users for adding as co-owner");
                }
                catch (Exception ex)
                {
                    _logWarn("Failed to load users: " + ex.Message);
                    MessageBox.Show("Failed to load users: " + ex.Message);
                }
            }

            private void BtnAdd_Click(object sender, EventArgs e)
            {
                var sel = cbUsers.SelectedItem as ListItem; if (sel == null) { MessageBox.Show("Select a user to add."); return; }
                btnAdd.Enabled = false; btnRemove.Enabled = false; cbUsers.Enabled = false; lbCoOwners.Enabled = false; var busy = new BusyForm("Adding co-owner..."); busy.Show(this);
                Task.Run(() =>
                {
                    try
                    {
                        var grant = new GrantAccessRequest { Target = new EntityReference("workflow", _flowId), PrincipalAccess = new PrincipalAccess { Principal = new EntityReference("systemuser", sel.Id), AccessMask = AccessRights.ReadAccess | AccessRights.WriteAccess | AccessRights.AppendAccess | AccessRights.AppendToAccess } };
                        _service.Execute(grant);
                        return (Exception)null;
                    }
                    catch (Exception ex) { return ex; }
                }).ContinueWith(t =>
                {
                    try { busy.Close(); } catch { }
                    btnAdd.Enabled = true; btnRemove.Enabled = true; cbUsers.Enabled = true; lbCoOwners.Enabled = true;
                    if (t.Result == null) { _logInfo($"Added co-owner {sel.Name}"); MessageBox.Show("Added co-owner."); }
                    else { _logWarn("Failed to add co-owner: " + t.Result.Message); MessageBox.Show("Failed to add co-owner: " + t.Result.Message); }
                    LoadData();
                }, TaskScheduler.FromCurrentSynchronizationContext());
            }

            private void BtnRemove_Click(object sender, EventArgs e)
            {
                var sel = lbCoOwners.SelectedItem as ListItem; if (sel == null) { MessageBox.Show("Select a co-owner to remove."); return; }
                var confirm = MessageBox.Show($"Are you sure you want to remove co-owner '{sel.Name}'?", "Confirm remove", MessageBoxButtons.YesNo, MessageBoxIcon.Question); if (confirm != DialogResult.Yes) return;
                btnAdd.Enabled = false; btnRemove.Enabled = false; cbUsers.Enabled = false; lbCoOwners.Enabled = false; var busy = new BusyForm("Removing co-owner..."); busy.Show(this);
                Task.Run(() =>
                {
                    try { var revoke = new RevokeAccessRequest { Target = new EntityReference("workflow", _flowId), Revokee = new EntityReference("systemuser", sel.Id) }; _service.Execute(revoke); return (Exception)null; }
                    catch (Exception ex) { return ex; }
                }).ContinueWith(t =>
                {
                    try { busy.Close(); } catch { }
                    btnAdd.Enabled = true; btnRemove.Enabled = true; cbUsers.Enabled = true; lbCoOwners.Enabled = true;
                    if (t.Result == null) { _logInfo($"Removed co-owner {sel.Name}"); MessageBox.Show("Removed co-owner."); }
                    else { _logWarn("Failed to remove co-owner: " + t.Result.Message); MessageBox.Show("Failed to remove co-owner: " + t.Result.Message); }
                    LoadData();
                }, TaskScheduler.FromCurrentSynchronizationContext());
            }

            private class ListItem { public Guid Id { get; set; } public string Name { get; set; } public override string ToString() => Name; }
        }

        private class ManageSolutionsDialog : Form
        {
            private IOrganizationService _service; private Guid _flowId; private Action<string> _logInfo; private Action<string> _logWarn;
            private ListBox lbSolutions; private ComboBox cbAvailableSolutions; private Button btnAddToSolution; private Button btnRemoveFromSolution; private Button btnClose;

            public ManageSolutionsDialog(IOrganizationService svc, Guid flowId, Action<string> logInfo = null, Action<string> logWarn = null) { _service = svc; _flowId = flowId; _logInfo = logInfo ?? (_ => { }); _logWarn = logWarn ?? (_ => { }); Initialize(); LoadData(); }

            private void Initialize()
            {
                this.Text = "Manage Flow Solutions"; this.Width = 700; this.Height = 420; this.StartPosition = FormStartPosition.CenterParent;
                lbSolutions = new ListBox() { Left = 10, Top = 10, Width = 420, Height = 320 }; this.Controls.Add(lbSolutions);
                cbAvailableSolutions = new ComboBox() { Left = 440, Top = 10, Width = 220, DropDownStyle = ComboBoxStyle.DropDownList }; this.Controls.Add(cbAvailableSolutions);
                btnAddToSolution = new Button() { Left = 440, Top = 50, Width = 220, Text = "Add to selected solution" }; btnAddToSolution.Click += BtnAddToSolution_Click; this.Controls.Add(btnAddToSolution);
                btnRemoveFromSolution = new Button() { Left = 440, Top = 90, Width = 220, Text = "Remove selected solution" }; btnRemoveFromSolution.Click += BtnRemoveFromSolution_Click; this.Controls.Add(btnRemoveFromSolution);
                btnClose = new Button() { Left = 440, Top = 320, Width = 220, Text = "Close" }; btnClose.Click += (s, e) => this.Close(); this.Controls.Add(btnClose);
                var lbl = new System.Windows.Forms.Label() { Left = 10, Top = 340, Width = 660, Text = "Note: Default Solution is excluded and cannot be added/removed." }; this.Controls.Add(lbl);
            }

            private void LoadData()
            {
                lbSolutions.Items.Clear(); cbAvailableSolutions.DataSource = null; cbAvailableSolutions.Items.Clear();
                try
                {
                    var scQ = new QueryExpression("solutioncomponent") { ColumnSet = new ColumnSet("solutionid", "objectid", "componenttype", "solutioncomponentid") };
                    scQ.Criteria.AddCondition("componenttype", ConditionOperator.Equal, 29);
                    scQ.Criteria.AddCondition("objectid", ConditionOperator.Equal, _flowId);
                    var scRes = _service.RetrieveMultiple(scQ);

                    Guid LocalGetGuidFromAttribute(Entity e, string attr) { if (e == null || string.IsNullOrEmpty(attr) || !e.Attributes.ContainsKey(attr)) return Guid.Empty; var val = e.Attributes[attr]; if (val == null) return Guid.Empty; if (val is Guid) return (Guid)val; if (val is EntityReference) return ((EntityReference)val).Id; if (val is AliasedValue) { var av = (AliasedValue)val; if (av.Value is Guid) return (Guid)av.Value; if (av.Value is EntityReference) return ((EntityReference)av.Value).Id; } return Guid.Empty; }

                    var solutionIds = scRes.Entities.Select(sc => LocalGetGuidFromAttribute(sc, "solutionid")).Where(g => g != Guid.Empty).Distinct().ToList();
                    var currentSolutions = new List<Tuple<Guid, string, string, bool>>();
                    if (solutionIds.Any())
                    {
                        var solQ = new QueryExpression("solution") { ColumnSet = new ColumnSet("solutionid", "friendlyname", "uniquename", "ismanaged") };
                        solQ.Criteria.AddCondition("solutionid", ConditionOperator.In, solutionIds.Cast<object>().ToArray());
                        var sols = _service.RetrieveMultiple(solQ);
                        foreach (var s in sols.Entities)
                        {
                            var id = s.Id; var friendly = s.GetAttributeValue<string>("friendlyname") ?? s.GetAttributeValue<string>("uniquename"); var uniq = s.GetAttributeValue<string>("uniquename") ?? string.Empty; if (!string.IsNullOrEmpty(uniq) && uniq.Equals("default", StringComparison.OrdinalIgnoreCase)) continue; if (!string.IsNullOrEmpty(friendly) && (friendly.IndexOf("default solution", StringComparison.OrdinalIgnoreCase) >= 0 || friendly.IndexOf("active solution", StringComparison.OrdinalIgnoreCase) >= 0)) continue; currentSolutions.Add(Tuple.Create(id, friendly, uniq, s.GetAttributeValue<bool?>("ismanaged") ?? false)); }
                    }

                    Func<Tuple<Guid, string, string, bool>, bool> isExcluded = t => { var uniq = t.Item3 ?? string.Empty; var friendly = t.Item2 ?? string.Empty; if (uniq.Equals("Default", StringComparison.OrdinalIgnoreCase) || uniq.Equals("default", StringComparison.OrdinalIgnoreCase)) return true; if (friendly.IndexOf("default solution", StringComparison.OrdinalIgnoreCase) >= 0) return true; if (friendly.IndexOf("active solution", StringComparison.OrdinalIgnoreCase) >= 0) return true; return false; };

                    foreach (var cs in currentSolutions.Where(t => !isExcluded(t)).OrderBy(t => t.Item2)) lbSolutions.Items.Add(new SolutionItem { Id = cs.Item1, FriendlyName = cs.Item2, UniqueName = cs.Item3, IsManaged = cs.Item4 });

                    var excludeIds = new HashSet<Guid>(currentSolutions.Select(t => t.Item1));
                    var availQ = new QueryExpression("solution") { ColumnSet = new ColumnSet("solutionid", "friendlyname", "uniquename", "ismanaged") };
                    availQ.Criteria.AddCondition("ismanaged", ConditionOperator.Equal, false);
                    var availRes = _service.RetrieveMultiple(availQ);
                    var availList = new List<SolutionItem>();
                    foreach (var s in availRes.Entities)
                    {
                        var id = s.Id; if (excludeIds.Contains(id)) continue; var friendly = s.GetAttributeValue<string>("friendlyname") ?? s.GetAttributeValue<string>("uniquename"); var uniq = s.GetAttributeValue<string>("uniquename"); var isManaged = s.GetAttributeValue<bool?>("ismanaged") ?? false; var item = new SolutionItem { Id = id, FriendlyName = friendly, UniqueName = uniq, IsManaged = isManaged }; if (isExcluded(Tuple.Create(id, friendly, uniq, isManaged))) continue; availList.Add(item);
                    }
                    availList = availList.OrderBy(x => x.FriendlyName).ToList(); cbAvailableSolutions.DisplayMember = "FriendlyName"; cbAvailableSolutions.ValueMember = "UniqueName"; cbAvailableSolutions.DataSource = availList; _logInfo($"Loaded {lbSolutions.Items.Count} current solutions and {availList.Count} available solutions");
                }
                catch (Exception ex) { _logWarn("Failed to load solutions: " + ex.Message); MessageBox.Show("Failed to load solutions: " + ex.Message); }
            }

            private void BtnAddToSolution_Click(object sender, EventArgs e)
            {
                var sel = cbAvailableSolutions.SelectedItem as SolutionItem; if (sel == null) { MessageBox.Show("Select a solution to add the flow to."); return; }
                var busy = new BusyForm("Adding flow to solution..."); busy.Show(this);
                Task.Run(() => { try { var req = new AddSolutionComponentRequest { SolutionUniqueName = sel.UniqueName, ComponentType = 29, ComponentId = _flowId }; _service.Execute(req); return (Exception)null; } catch (Exception ex) { return ex; } }).ContinueWith(t => { try { busy.Close(); } catch { } if (t.Result == null) { MessageBox.Show("Flow added to solution."); } else { MessageBox.Show("Failed to add flow to solution: " + t.Result.Message); } LoadData(); }, TaskScheduler.FromCurrentSynchronizationContext());
            }

            private void BtnRemoveFromSolution_Click(object sender, EventArgs e)
            {
                var sel = lbSolutions.SelectedItem as SolutionItem; if (sel == null) { MessageBox.Show("Select a solution to remove the flow from."); return; }
                var confirm = MessageBox.Show($"Remove flow from solution '{sel.FriendlyName}'?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question); if (confirm != DialogResult.Yes) return;
                var busy = new BusyForm("Removing flow from solution..."); busy.Show(this);
                Task.Run(() => { try { var req = new RemoveSolutionComponentRequest { SolutionUniqueName = sel.UniqueName, ComponentType = 29, ComponentId = _flowId }; _service.Execute(req); return (Exception)null; } catch (Exception ex) { return ex; } }).ContinueWith(t => { try { busy.Close(); } catch { } if (t.Result == null) { MessageBox.Show("Flow removed from solution."); } else { MessageBox.Show("Failed to remove flow from solution: " + t.Result.Message); } LoadData(); }, TaskScheduler.FromCurrentSynchronizationContext());
            }

            private class SolutionItem { public Guid Id { get; set; } public string FriendlyName { get; set; } public string UniqueName { get; set; } public bool IsManaged { get; set; } public override string ToString() => FriendlyName; }
        }

        private bool TryOpenConnectionDialog()
        {
            try
            {
                // XrmToolBox exposes a ConnectionManager via static classes or services; attempt common approaches via reflection
                var asm = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(a => a.GetName().Name == "XrmToolBox");
                if (asm != null)
                {
                    var mainFormType = asm.GetType("XrmToolBox.MainForm");
                    if (mainFormType != null)
                    {
                        // find an instance of MainForm
                        var mainForm = Application.OpenForms.Cast<Form>().FirstOrDefault(f => f.GetType() == mainFormType);
                        if (mainForm != null)
                        {
                            // try method ShowConnectionDialog or OpenConnectionManager or Connect
                            var m = mainFormType.GetMethod("ShowConnectionDialog", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
                            if (m != null) { m.Invoke(mainForm, null); return true; }
                            m = mainFormType.GetMethod("OpenConnectionManager", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
                            if (m != null) { m.Invoke(mainForm, null); return true; }
                        }
                    }
                }

                // fallback: try to find any open form with 'Connect' menu item and invoke click
                foreach (Form f in Application.OpenForms)
                {
                    var mi = f.GetType().GetMethod("OpenConnectionDialog", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
                    if (mi != null) { mi.Invoke(f, null); return true; }
                }
            }
            catch { }
            return false;
        }

        private void toolStripMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
    }
}