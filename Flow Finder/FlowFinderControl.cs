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
using System.ServiceModel;


namespace Flow_Finder
{
    public partial class FlowFinderControl : PluginControlBase, IGitHubPlugin
    {
        // Implement IGitHubPlugin so XrmToolBox will show the standard feedback -> New issue menu
        internal static readonly string GitHubRepoName = "FlowFinder";
        internal static readonly string GitHubUserName = "MattCollins-Jones";
        public string RepositoryName => GitHubRepoName;
        public string UserName => GitHubUserName;

        private Settings mySettings;
        private DataTable flowsTable;
        private List<FlowInfo> lastResults = new List<FlowInfo>();
        private Dictionary<string, bool> _solutionManagedStatus = new Dictionary<string, bool>();
        private Dictionary<Guid, string> _flowClientData = new Dictionary<Guid, string>();
        private Font _boldFont;

        internal const int WorkflowCategoryCloudFlow = 6;
        internal const int WorkflowCategoryScheduledFlow = 5;
        internal const int WorkflowCategoryDesktopFlow = 7;
        internal const int SolutionComponentTypeCloudFlow = 29;
        internal const int WorkflowEntityTypeCode = 4703;

        private class FlowInfo
        {
            public Guid Id { get; set; }
            public string Name { get; set; }
            public bool InSolution { get; set; }
            public bool IsInManagedSolution { get; set; }
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
            public string Status { get; set; }
            public string LinkToFlow { get; set; }
            public string ClientDataJson { get; set; }
        }

        private string MapStateCodeToStatus(int stateCode)
        {
            switch (stateCode)
            {
                case 0:
                    return "Inactive";
                case 1:
                    return "Active";
                case 2:
                    return "Suspended";
                default:
                    return "Unknown";
            }
        }

        public FlowFinderControl()
        {
            InitializeComponent();
            InitializeFlowsTable();
            dgvFlows.DataBindingComplete += DgvFlows_DataBindingComplete;
            this.Disposed += (s, e) => { _boldFont?.Dispose(); _boldFont = null; };
        }

        private void DgvFlows_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewRow row in dgvFlows.Rows)
            {
                if (row.IsNewRow) continue;
                var rowView = row.DataBoundItem as DataRowView;
                if (rowView == null) continue;
                if (rowView["FlowId"] is Guid id && id != Guid.Empty)
                    row.Tag = id;
            }
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
            flowsTable.Columns.Add("Triggering Table"); // Updated from "Triggering Entity"
            flowsTable.Columns.Add("Other Data Sources");
            flowsTable.Columns.Add("Link to Flow");
            flowsTable.Columns.Add("Status");
            // hidden flag column used for filtering (added at end so existing column indexes remain stable)
            flowsTable.Columns.Add("IsInManagedSolution", typeof(bool));
            flowsTable.Columns.Add("FlowId", typeof(Guid));

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

        // Names to exclude from disabled checks (case-insensitive)
        internal static readonly HashSet<string> DisabledUserExceptions = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "System" };

        private bool IsUserDisabled(Guid userId, out string fullName)
        {
            fullName = null;
            if (userId == Guid.Empty) return false;
            try
            {
                var u = Service.Retrieve("systemuser", userId, new ColumnSet("systemuserid", "fullname", "isdisabled"));
                if (u == null) return false;
                fullName = GetStringSafe(u, "fullname");
                // exclude known system/test accounts by name (trim whitespace)
                if (!string.IsNullOrEmpty(fullName) && DisabledUserExceptions.Contains(fullName.Trim())) return false;
                try { if (u.Contains("isdisabled") && u.GetAttributeValue<bool>("isdisabled")) return true; } catch { }
            }
            catch { }
            return false;
        }

        private string RemoveDisabledMarker(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            const string marker = "(disabled)";
            var idx = s.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (idx >= 0) s = s.Remove(idx, marker.Length);
            return s.Trim();
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
            catch (Exception ex)
            {
                LogWarning($"Failed to parse client data: {ex.Message}");
            }
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

            // legacy single-purpose filter preserved but redirect to ApplyFilters
            ApplyFilters();
        }

        private void ApplyFilters()
        {
            if (flowsTable == null) return;
            var filters = new List<string>();
            // solution dropdown
            try
            {
                if (cmbSolutions != null && cmbSolutions.SelectedIndex > 0)
                {
                    var sel = cmbSolutions.ComboBox.Items[cmbSolutions.SelectedIndex];
                    if (sel != null)
                    {
                        var solutionName = sel.ToString().Replace("'", "''");
                        filters.Add($"Solutions LIKE '%{solutionName}%'");
                    }
                }
            }
            catch { }

            // hide managed toggle
            try
            {
                if (chkHideManaged != null && chkHideManaged.Checked)
                {
                    filters.Add("IsInManagedSolution = false");
                }
            }
            catch { }

            var dv = flowsTable.DefaultView;
            dv.RowFilter = filters.Any() ? string.Join(" AND ", filters) : string.Empty;
            // Reapply conditional formatting after filtering
            try { ApplyConditionalFormattingToFlowsGrid(); } catch { }
        }

        // Reapply conditional formatting to rows (disabled owner/Co-Owner highlighting).
        // Owner/Co-Owner "(disabled)" markers are resolved on the background thread — no network calls here.
        private void ApplyConditionalFormattingToFlowsGrid()
        {
            if (dgvFlows == null || dgvFlows.Rows.Count == 0 || lastResults == null || lastResults.Count == 0) return;

            if (_boldFont == null || _boldFont.Size != dgvFlows.Font.Size || _boldFont.FontFamily.Name != dgvFlows.Font.FontFamily.Name)
            {
                _boldFont?.Dispose();
                _boldFont = new Font(dgvFlows.Font, FontStyle.Bold);
            }

            var byId = lastResults.ToDictionary(f => f.Id);

            for (int i = 0; i < dgvFlows.Rows.Count; i++)
            {
                var row = dgvFlows.Rows[i];
                if (row.IsNewRow) continue;
                try
                {
                    var rowView = row.DataBoundItem as DataRowView;
                    if (rowView == null) continue;

                    FlowInfo fi = null;
                    if (rowView["FlowId"] is Guid id && id != Guid.Empty)
                    {
                        byId.TryGetValue(id, out fi);
                        row.Tag = id;
                    }
                    if (fi == null) continue;

                    bool hasDisabled = false;
                    if (!string.IsNullOrEmpty(fi.Owner) && fi.Owner.IndexOf("(disabled)", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        var ownerBase = RemoveDisabledMarker(fi.Owner);
                        if (!DisabledUserExceptions.Contains(ownerBase)) hasDisabled = true;
                    }
                    if (!hasDisabled && !string.IsNullOrEmpty(fi.CoOwners) && fi.CoOwners.IndexOf("(disabled)", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        foreach (var p in fi.CoOwners.Split(','))
                        {
                            if (p.IndexOf("(disabled)", StringComparison.OrdinalIgnoreCase) >= 0 &&
                                !DisabledUserExceptions.Contains(RemoveDisabledMarker(p)))
                            {
                                hasDisabled = true;
                                break;
                            }
                        }
                    }

                    if (hasDisabled)
                    {
                        row.DefaultCellStyle.BackColor = Color.LightYellow;
                        row.DefaultCellStyle.SelectionBackColor = Color.Gold;
                        row.DefaultCellStyle.ForeColor = Color.DarkRed;
                        row.DefaultCellStyle.Font = _boldFont;
                    }
                    else
                    {
                        row.DefaultCellStyle.BackColor = Color.Empty;
                        row.DefaultCellStyle.SelectionBackColor = Color.Empty;
                        row.DefaultCellStyle.ForeColor = Color.Empty;
                        row.DefaultCellStyle.Font = dgvFlows.Font;
                    }
                }
                catch { }
            }
        }

        private void chkHideManaged_Click(object sender, EventArgs e)
        {
            try
            {
                // Update label to indicate the action that pressing the button will perform
                try { if (chkHideManaged != null) chkHideManaged.Text = chkHideManaged.Checked ? "Show managed" : "Hide managed"; } catch { }

                // Apply filters to the flows table
                ApplyFilters();

                // Update the solution dropdown filter to show/hide managed solutions
                if (cmbSolutions != null && cmbSolutions.ComboBox != null)
                {
                    var allSolutions = _solutionManagedStatus.Keys.OrderBy(s => s).ToList();
                    cmbSolutions.ComboBox.Items.Clear();
                    cmbSolutions.ComboBox.Items.Add("All solutions");

                    if (chkHideManaged.Checked)
                    {
                        // Exclude managed solutions
                        foreach (var solution in allSolutions)
                        {
                            if (!_solutionManagedStatus.ContainsKey(solution) || !_solutionManagedStatus[solution])
                            {
                                cmbSolutions.ComboBox.Items.Add(solution);
                            }
                        }
                    }
                    else
                    {
                        // Include all solutions
                        foreach (var solution in allSolutions)
                        {
                            cmbSolutions.ComboBox.Items.Add(solution);
                        }
                    }

                    cmbSolutions.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                LogWarning($"Error in chkHideManaged_Click: {ex.Message}");
            }
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
            catch (Exception ex) { LogWarning("Connection check failed: " + ex.Message); }
             WorkAsync(new WorkAsyncInfo
             {
                Message = "Retrieving cloud flows and solutions",
                Work = (worker, args) =>
                {
                    LogInfo("Starting retrieval of cloud flows and solutions");

                    var environmentId = ConnectionDetail.EnvironmentId?.ToString() ?? string.Empty;
                    Guid defaultSolutionId = Guid.Empty;
                    try
                    {
                        var defaultSolQuery = new QueryExpression("solution")
                        {
                            ColumnSet = new ColumnSet("solutionid"),
                            Criteria = new FilterExpression()
                        };
                        defaultSolQuery.Criteria.AddCondition("uniquename", ConditionOperator.Equal, "Default");
                        var defaultSolResult = Service.RetrieveMultiple(defaultSolQuery);
                        if (defaultSolResult.Entities.Any())
                        {
                            defaultSolutionId = defaultSolResult.Entities.First().Id;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogWarning("Failed to retrieve default solution ID: " + ex.Message);
                    }

                    var qe = new QueryExpression("workflow") { ColumnSet = new ColumnSet("workflowid", "name", "ownerid", "type", "category", "description", "createdby", "clientdata", "statecode"), Criteria = new FilterExpression() };
                    qe.Criteria.AddCondition("type", ConditionOperator.Equal, 1);
                    qe.Criteria.AddCondition("category", ConditionOperator.Equal, WorkflowCategoryCloudFlow);
                    var flowEntities = RetrieveAllPages(Service, qe);

                    if (flowEntities.Count == 0)
                    {
                        var qe2 = new QueryExpression("workflow") { ColumnSet = new ColumnSet("workflowid", "name", "ownerid", "type", "category", "description", "createdby", "clientdata", "statecode"), Criteria = new FilterExpression() };
                        qe2.Criteria.AddCondition("type", ConditionOperator.Equal, 1);
                        qe2.Criteria.AddCondition("category", ConditionOperator.In, new object[] { WorkflowCategoryScheduledFlow, WorkflowCategoryCloudFlow, WorkflowCategoryDesktopFlow });
                        flowEntities = RetrieveAllPages(Service, qe2);
                    }
                    if (flowEntities.Count == 0)
                    {
                        var qe3 = new QueryExpression("workflow") { ColumnSet = new ColumnSet("workflowid", "name", "ownerid", "type", "category", "description", "createdby", "clientdata", "statecode") };
                        qe3.Criteria.AddCondition("type", ConditionOperator.Equal, 1);
                        flowEntities = RetrieveAllPages(Service, qe3);
                    }

                    var scQ = new QueryExpression("solutioncomponent") { ColumnSet = new ColumnSet("solutionid", "objectid", "componenttype"), Criteria = new FilterExpression() };
                    scQ.Criteria.AddCondition("componenttype", ConditionOperator.Equal, SolutionComponentTypeCloudFlow);
                    var solutionComponentEntities = RetrieveAllPages(Service, scQ);

                    var solutionIds = solutionComponentEntities.Select(sc => GetGuidFromAttribute(sc, "solutionid")).Where(g => g != Guid.Empty).Distinct().ToList();
                    var solutionNames = new Dictionary<Guid, string>();
                    var solutionIsManaged = new Dictionary<Guid, bool>();
                    // built locally on the background thread and assigned to the field only in PostWorkCallBack (UI thread) — no data race
                    var newManagedStatus = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
                    if (solutionIds.Any())
                    {
                        foreach (var idBatch in ToBatches(solutionIds, 500))
                        {
                            try
                            {
                                var solFetch = new QueryExpression("solution") { ColumnSet = new ColumnSet("solutionid", "friendlyname", "uniquename", "ismanaged") };
                                solFetch.Criteria.AddCondition("solutionid", ConditionOperator.In, idBatch.Cast<object>().ToArray());
                                foreach (var s in RetrieveAllPages(Service, solFetch))
                                {
                                    var friendly = GetStringSafe(s, "friendlyname") ?? GetStringSafe(s, "uniquename");
                                    var uniq = GetStringSafe(s, "uniquename") ?? string.Empty;
                                    if (!string.IsNullOrEmpty(uniq) && uniq.Equals("default", StringComparison.OrdinalIgnoreCase)) continue;
                                    if (!string.IsNullOrEmpty(friendly) && (friendly.IndexOf("default solution", StringComparison.OrdinalIgnoreCase) >= 0 || friendly.IndexOf("active solution", StringComparison.OrdinalIgnoreCase) >= 0)) continue;
                                    solutionNames[s.Id] = friendly;
                                    bool isManaged = false;
                                    try { isManaged = s.GetAttributeValue<bool?>("ismanaged") ?? false; } catch { }
                                    solutionIsManaged[s.Id] = isManaged;
                                    if (!string.IsNullOrEmpty(friendly) && !newManagedStatus.ContainsKey(friendly))
                                    {
                                        newManagedStatus.Add(friendly, isManaged);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                LogWarning($"Failed to fetch solution names for batch: {ex.Message}");
                            }
                        }
                    }

                    var flowSolutionNames = new Dictionary<Guid, List<string>>();
                    var flowSolutionIds = new Dictionary<Guid, List<Guid>>();
                    foreach (var sc in solutionComponentEntities)
                    {
                        var objId = GetGuidFromAttribute(sc, "objectid");
                        var solId = GetGuidFromAttribute(sc, "solutionid");
                        if (objId == Guid.Empty || solId == Guid.Empty) continue;
                        if (!flowSolutionNames.ContainsKey(objId)) flowSolutionNames[objId] = new List<string>();
                        if (!flowSolutionIds.ContainsKey(objId)) flowSolutionIds[objId] = new List<Guid>();
                        if (solutionNames.ContainsKey(solId)) flowSolutionNames[objId].Add(solutionNames[solId]);
                        flowSolutionIds[objId].Add(solId);
                    }

                    var allSolutionNames = flowSolutionNames.SelectMany(kvp => kvp.Value).Distinct().OrderBy(s => s).ToList();

                    var results = new List<FlowInfo>();
                    var principalIds = new HashSet<Guid>();
                    var ownerIds = new HashSet<Guid>();

                    // Attempt a single bulk query of principalobjectaccess for all flows at once.
                    // Falls back to per-flow RetrieveSharedPrincipalsAndAccessRequest if the caller lacks prvReadPOA.
                    var flowPrincipals = new Dictionary<Guid, List<Guid>>();
                    bool poaQuerySucceeded = false;
                    try
                    {
                        var flowIds = flowEntities.Select(f => f.Id).ToList();
                        int totalPoa = 0;
                        foreach (var poaBatch in ToBatches(flowIds, 500))
                        {
                            var poaQ = new QueryExpression("principalobjectaccess")
                            {
                                ColumnSet = new ColumnSet("objectid", "principalid", "accessrightsmask")
                            };
                            poaQ.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, WorkflowEntityTypeCode);
                            poaQ.Criteria.AddCondition("objectid", ConditionOperator.In, poaBatch.Cast<object>().ToArray());
                            // Only include records with a direct (non-inherited) share — accessrightsmask > 0.
                            // Records with only team-inherited access have accessrightsmask = 0 and would otherwise
                            // cause revoked co-owners to reappear on full reload.
                            poaQ.Criteria.AddCondition("accessrightsmask", ConditionOperator.GreaterThan, 0);
                            foreach (var poa in RetrieveAllPages(Service, poaQ))
                            {
                                totalPoa++;
                                var objId = GetGuidFromAttribute(poa, "objectid");
                                var principalId = GetGuidFromAttribute(poa, "principalid");
                                if (objId == Guid.Empty || principalId == Guid.Empty) continue;
                                if (!flowPrincipals.ContainsKey(objId)) flowPrincipals[objId] = new List<Guid>();
                                flowPrincipals[objId].Add(principalId);
                            }
                        }
                        poaQuerySucceeded = true;
                        LogInfo($"POA bulk query returned {totalPoa} records for {flowIds.Count} flows.");
                    }
                    catch (FaultException<OrganizationServiceFault> fex) when (fex.Detail?.ErrorCode == -2147220960)
                    {
                        LogWarning("Insufficient privilege to query principalobjectaccess — falling back to per-flow share lookup.");
                    }
                    catch (Exception ex)
                    {
                        LogWarning($"POA bulk query failed ({ex.Message}) — falling back to per-flow share lookup.");
                    }

                    foreach (var f in flowEntities)
                    {
                        var stateCode = f.Contains("statecode") ? f.GetAttributeValue<OptionSetValue>("statecode").Value : -1;
                        var flowId = f.Id;
                        var solutionIdForLink = flowSolutionIds.ContainsKey(flowId) && flowSolutionIds[flowId].Any()
                            ? flowSolutionIds[flowId].First()
                            : defaultSolutionId;

                        var fi = new FlowInfo
                        {
                            Id = flowId,
                            Name = GetStringSafe(f, "name"),
                            Owner = GetEntityReferenceName(f, "ownerid"),
                            OwnerId = GetGuidFromAttribute(f, "ownerid"),
                            Description = GetStringSafe(f, "description"),
                            CreatedBy = GetEntityReferenceName(f, "createdby"),
                            Status = MapStateCodeToStatus(stateCode),
                            LinkToFlow = solutionIdForLink != Guid.Empty && !string.IsNullOrEmpty(environmentId)
                                ? $"https://make.powerautomate.com/environments/{environmentId}/solutions/{solutionIdForLink}/flows/{flowId}?v3=true"
                                : string.Empty
                        };

                        if (flowSolutionNames.ContainsKey(f.Id)) { fi.InSolution = true; fi.Solutions = string.Join(", ", flowSolutionNames[f.Id].Distinct()); }
                        else fi.InSolution = false;
                        // mark managed state if any containing solution is managed
                        try { fi.IsInManagedSolution = flowSolutionIds.ContainsKey(f.Id) && flowSolutionIds[f.Id].Any(sid => solutionIsManaged.ContainsKey(sid) && solutionIsManaged[sid]); } catch { fi.IsInManagedSolution = false; }

                        if (poaQuerySucceeded)
                        {
                            var principalGuids = flowPrincipals.ContainsKey(fi.Id)
                                ? flowPrincipals[fi.Id].Distinct().ToList()
                                : new List<Guid>();
                            fi.PrincipalIds = principalGuids;
                            foreach (var id in principalGuids) principalIds.Add(id);
                        }
                        else
                        {
                            try
                            {
                                var req = new RetrieveSharedPrincipalsAndAccessRequest { Target = new EntityReference("workflow", f.Id) };
                                var resp = (RetrieveSharedPrincipalsAndAccessResponse)Service.Execute(req);
                                var principalGuids = new List<Guid>();
                                if (resp?.PrincipalAccesses != null)
                                {
                                    foreach (var pa in resp.PrincipalAccesses)
                                    {
                                        if (pa.Principal is EntityReference principalRef)
                                        {
                                            principalGuids.Add(principalRef.Id);
                                            principalIds.Add(principalRef.Id);
                                        }
                                    }
                                }
                                fi.PrincipalIds = principalGuids;
                            }
                            catch (Exception ex) { LogWarning($"Failed to retrieve shared principals for flow {fi.Id}: {ex.Message}"); }
                        }

                        // track owner ids for later disabled check
                        if (fi.OwnerId != Guid.Empty) ownerIds.Add(fi.OwnerId);

                        // parse clientdata fetched in the initial query — avoids a separate Retrieve per flow
                        try
                        {
                            var clientJson = GetStringSafe(f, "clientdata");
                            if (!string.IsNullOrEmpty(clientJson))
                            {
                                fi.ClientDataJson = clientJson;
                                ParseClientData(clientJson, fi);
                            }
                        }
                        catch (Exception ex)
                        {
                            LogWarning($"Failed to parse clientdata for flow {fi.Id}: {ex.Message}");
                        }

                        results.Add(fi);
                    }

                    var principalNames = new Dictionary<Guid, string>();
                    var principalDisabled = new HashSet<Guid>();
                    // include owner ids alongside shared principals so we can detect disabled primary owners too
                    var userIdsToCheck = new HashSet<Guid>(principalIds);
                    foreach (var oid in ownerIds) userIdsToCheck.Add(oid);
                    if (userIdsToCheck.Any())
                    {
                        var idsArray = userIdsToCheck.ToArray();
                        var userQ = new QueryExpression("systemuser") { ColumnSet = new ColumnSet("systemuserid", "fullname", "isdisabled") };
                        userQ.Criteria.AddCondition("systemuserid", ConditionOperator.In, idsArray.Cast<object>().ToArray());
                        var users = Service.RetrieveMultiple(userQ);
                        foreach (var u in users.Entities)
                        {
                            var fullname = GetStringSafe(u, "fullname");
                            principalNames[u.Id] = fullname;
                            bool isDisabled = false;
                            try { if (u.Contains("isdisabled")) isDisabled = u.GetAttributeValue<bool>("isdisabled"); } catch { }
                            // Do not treat listed exception names as disabled (trim whitespace)
                            try { if (!string.IsNullOrEmpty(fullname) && DisabledUserExceptions.Contains(fullname.Trim())) isDisabled = false; } catch { }
                            if (isDisabled) principalDisabled.Add(u.Id);
                        }

                        var teamQ = new QueryExpression("team") { ColumnSet = new ColumnSet("teamid", "name") };
                        teamQ.Criteria.AddCondition("teamid", ConditionOperator.In, idsArray.Cast<object>().ToArray());
                        var teams = Service.RetrieveMultiple(teamQ);
                        foreach (var t in teams.Entities) principalNames[t.Id] = GetStringSafe(t, "name");
                    }

                    // As a fallback, ensure primary owners are resolved even if not returned above
                    try
                    {
                        foreach (var fi in results)
                        {
                            if (fi.OwnerId != Guid.Empty && !principalNames.ContainsKey(fi.OwnerId))
                            {
                                try
                                {
                                    var ownerEnt = Service.Retrieve("systemuser", fi.OwnerId, new ColumnSet("systemuserid", "fullname", "isdisabled"));
                                    if (ownerEnt != null && ownerEnt.Id != Guid.Empty)
                                    {
                                        var ownerFull = GetStringSafe(ownerEnt, "fullname");
                                        principalNames[ownerEnt.Id] = ownerFull;
                                        try
                                        {
                                            bool ownerIsDisabled = false;
                                            try { ownerIsDisabled = ownerEnt.Contains("isdisabled") && ownerEnt.GetAttributeValue<bool>("isdisabled"); } catch { }
                                            // Do not treat listed exception names as disabled
                                            if (!string.IsNullOrEmpty(ownerFull) && DisabledUserExceptions.Contains(ownerFull.Trim())) ownerIsDisabled = false;
                                            if (ownerIsDisabled) principalDisabled.Add(ownerEnt.Id);
                                        }
                                        catch { }
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                    catch { }

                    foreach (var fi in results)
                    {
                        if (fi.PrincipalIds != null && fi.PrincipalIds.Any())
                        {
                            var coIds = fi.PrincipalIds.Where(id => id != fi.OwnerId).Distinct().ToList();
                            var coNames = new List<string>();
                            foreach (var id in coIds)
                            {
                                var name = principalNames.ContainsKey(id) ? principalNames[id] : id.ToString();
                                if (principalDisabled.Contains(id)) name += " (disabled)";
                                coNames.Add(name);
                            }
                             fi.CoOwners = coNames.Any() ? string.Join(", ", coNames) : null;
                        }
                        else fi.CoOwners = null;

                        // prefer resolved name for primary owner and annotate if disabled
                        try
                        {
                            if (fi.OwnerId != Guid.Empty && principalNames.ContainsKey(fi.OwnerId)) fi.Owner = principalNames[fi.OwnerId];
                            if (fi.OwnerId != Guid.Empty && principalDisabled.Contains(fi.OwnerId)) fi.Owner = (fi.Owner ?? string.Empty) + " (disabled)";
                        }
                        catch { }
                    }

                    args.Result = new FindFlowsResult { Results = results, Solutions = allSolutionNames, DisabledPrincipalIds = principalDisabled.ToArray(), ManagedStatus = newManagedStatus };
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        MessageBox.Show(args.Error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    var r = (FindFlowsResult)args.Result;
                    lastResults = r.Results.OrderBy(f => f.Name).ToList();
                    var solutionNames = r.Solutions;
                    var disabledIds = r.DisabledPrincipalIds;
                    _solutionManagedStatus = r.ManagedStatus;

                    _flowClientData.Clear();
                    foreach (var f in lastResults)
                        if (!string.IsNullOrEmpty(f.ClientDataJson))
                            _flowClientData[f.Id] = f.ClientDataJson;

                    try { dgvFlows.DataSource = null; } catch { }
                    flowsTable.Clear();
                    dgvFlows.SuspendLayout();
                    try
                    {
                        foreach (var f in lastResults)
                        {
                            var row = flowsTable.NewRow();
                            row["Name"] = f.Name ?? string.Empty;
                            row["Status"] = f.Status ?? string.Empty;
                            row["Link to Flow"] = f.LinkToFlow ?? string.Empty;
                            row["Description"] = f.Description ?? string.Empty;
                            row["Solutions"] = f.Solutions ?? string.Empty;
                            row["Primary Owner"] = f.Owner ?? string.Empty;
                            row["Co-Owners"] = f.CoOwners ?? string.Empty;
                            row["Triggering Source"] = f.TriggerSource ?? string.Empty;
                            row["Triggering Table"] = f.TriggerEntity ?? string.Empty;
                            row["Other Data Sources"] = f.OtherDataSources ?? string.Empty;
                            row["IsInManagedSolution"] = f.IsInManagedSolution;
                            row["FlowId"] = f.Id;
                            flowsTable.Rows.Add(row);
                        }
                        dgvFlows.DataSource = flowsTable;
                    }
                    finally
                    {
                        dgvFlows.ResumeLayout(false);
                    }
                    // hide the helper columns
                    try { if (dgvFlows.Columns.Contains("IsInManagedSolution")) dgvFlows.Columns["IsInManagedSolution"].Visible = false; } catch { }
                    try { if (dgvFlows.Columns.Contains("FlowId")) dgvFlows.Columns["FlowId"].Visible = false; } catch { }
                    // ApplyFilters calls ApplyConditionalFormattingToFlowsGrid — no need to call it again afterwards
                    try { ApplyFilters(); } catch { }

                    try { cmbSolutions.ComboBox.DataSource = null; } catch { }
                    cmbSolutions.ComboBox.Items.Clear(); cmbSolutions.ComboBox.Items.Add("All solutions");
                    foreach (var sname in solutionNames) cmbSolutions.ComboBox.Items.Add(sname);
                    try { cmbSolutions.SelectedIndex = 0; } catch { cmbSolutions.ComboBox.SelectedIndex = 0; }
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
                    AddRuntimeButtonIfMissing("tsbManageCoOwners", "Manage Co-Owners", global::FlowFinder.Properties.Resources.UsersMan, tsbManageCoOwners_Click);
                    AddRuntimeButtonIfMissing("tsbManageSolutions", "Manage Solutions", global::FlowFinder.Properties.Resources.SolIcon, tsbManageSolutions_Click);
                    AddRuntimeButtonIfMissing("tsbViewSchema", "View Schema", CreateCodeIcon(), tsbViewSchema_Click);
                    AddRuntimeButtonIfMissing("tsbSettingsRuntime", "Settings", global::FlowFinder.Properties.Resources.Settings, tsbSettings_Click);
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
                    try { cmbSolutions.SelectedIndexChanged += new EventHandler((s,ev) => { try { if (cmbSolutions.SelectedIndex >= 0) ApplyFilters(); } catch { } }); } catch { }
                    // initialize hide-managed button text to reflect current checked state
                    try { if (chkHideManaged != null) chkHideManaged.Text = chkHideManaged.Checked ? "Show managed" : "Hide managed"; } catch { }
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

        private void AddRuntimeButtonIfMissing(string name, string text, Image image, EventHandler handler)
        {
            bool has = false;
            foreach (ToolStripItem it in toolStripMenu.Items)
            {
                if (string.Equals(it.Name, name, StringComparison.OrdinalIgnoreCase)) { has = true; break; }
            }
            if (!has)
            {
                var btn = new ToolStripButton() { DisplayStyle = ToolStripItemDisplayStyle.ImageAndText, Name = name, Text = text, Image = image };
                btn.Click += handler;
                toolStripMenu.Items.Add(new ToolStripSeparator());
                toolStripMenu.Items.Add(btn);
            }
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

        private static Image CreateCodeIcon()
        {
            const int size = 16;
            var bmp = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                using (var font = new Font("Segoe UI", 6.5f, FontStyle.Bold, GraphicsUnit.Point))
                using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
                {
                    var text = "</>";
                    var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                    g.DrawString(text, font, brush, new RectangleF(0, 0, size, size), sf);
                }
            }
            return bmp;
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
                    writer.WriteLine("Name,Description,Solutions,Primary Owner,Co-Owners,Triggering Source,Triggering Table,Other Data Sources,Link to Flow,Status");
                    foreach (DataGridViewRow row in dgvFlows.Rows)
                    {
                        if (row.IsNewRow) continue;
                        var vals = new List<string>();
                        for (int i = 0; i < 10; i++) vals.Add(EscapeCsv(Convert.ToString(row.Cells[i].Value ?? string.Empty)));
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

        private sealed class FindFlowsResult
        {
            public List<FlowInfo> Results { get; set; }
            public List<string> Solutions { get; set; }
            public Guid[] DisabledPrincipalIds { get; set; }
            public Dictionary<string, bool> ManagedStatus { get; set; }
        }

        private void tsbManageCoOwners_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = null;
            if (dgvFlows.SelectedRows != null && dgvFlows.SelectedRows.Count > 0) row = dgvFlows.SelectedRows[0];
            else row = dgvFlows.CurrentRow;
            if (row == null) { MessageBox.Show("Please select a flow in the list first.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (row.Tag == null || !(row.Tag is Guid flowId)) return;
            using (var dlg = new ManageCoOwnersDialog(Service, flowId, s => LogInfo(s), s => LogWarning(s)))
            {
                dlg.ShowDialog();
                try
                {
                    // Wait briefly for any in-flight add/remove operation to complete so refresh sees changes
                    dlg.LastOperation?.Wait(5000);
                }
                catch { }

                // Only refresh if the dialog made changes
                if (dlg.ChangesMade)
                {
                    RefreshAccordingtoSetting((Guid)row.Tag);
                }
            }
        }

        private void tsbManageSolutions_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = null;
            if (dgvFlows.SelectedRows != null && dgvFlows.SelectedRows.Count > 0) row = dgvFlows.SelectedRows[0];
            else row = dgvFlows.CurrentRow;
            if (row == null) { MessageBox.Show("Please select a flow in the list first.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (row.Tag == null || !(row.Tag is Guid flowId)) return;
            try
            {
                using (var dlg = new ManageSolutionsDialog(Service, flowId, s => LogInfo(s), s => LogWarning(s)))
                {
                    dlg.ShowDialog();
                    try { dlg.LastOperation?.Wait(5000); } catch { }
                    if (dlg.ChangesMade) RefreshAccordingtoSetting((Guid)row.Tag);
                }
            }
            catch (Exception ex)
            {
                LogWarning("Manage solutions dialog failed: " + ex.Message);
                try { MessageBox.Show("Failed to open Manage Solutions: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
            }
        }

        private void tsbViewSchema_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = null;
            if (dgvFlows.SelectedRows != null && dgvFlows.SelectedRows.Count > 0) row = dgvFlows.SelectedRows[0];
            else row = dgvFlows.CurrentRow;
            if (row == null) { MessageBox.Show("Please select a flow in the list first.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (row.Tag == null || !(row.Tag is Guid flowId)) return;

            if (!_flowClientData.TryGetValue(flowId, out var json) || string.IsNullOrEmpty(json))
            {
                MessageBox.Show("No schema (clientdata) is available for the selected flow.", "View Schema", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var flowName = row.Cells["Name"].Value?.ToString() ?? flowId.ToString();
            using (var dlg = new ViewSchemaDialog(flowName, json))
                dlg.ShowDialog(this);
        }

        private void RefreshAccordingtoSetting(Guid flowId)
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
                    var disabledSet = new HashSet<Guid>();
                    try
                    {
                        var wf = Service.Retrieve("workflow", flowId, new ColumnSet("workflowid", "name", "ownerid", "description", "createdby", "clientdata", "statecode"));
                        var stateCode = wf.Contains("statecode") ? wf.GetAttributeValue<OptionSetValue>("statecode").Value : -1;
                        fi.Id = wf.Id; fi.Name = GetStringSafe(wf, "name"); fi.Description = GetStringSafe(wf, "description"); fi.Owner = GetEntityReferenceName(wf, "ownerid"); fi.OwnerId = GetGuidFromAttribute(wf, "ownerid"); fi.CreatedBy = GetEntityReferenceName(wf, "createdby");
                        fi.Status = MapStateCodeToStatus(stateCode);
                        try { var cj = GetStringSafe(wf, "clientdata"); if (!string.IsNullOrEmpty(cj)) fi.ClientDataJson = cj; } catch { }

                        // --- Start of fix: Fetch solution info for the single flow ---
                        var environmentId = ConnectionDetail.EnvironmentId?.ToString() ?? string.Empty;
                        Guid defaultSolutionId = Guid.Empty;
                        try
                        {
                            var defaultSolQuery = new QueryExpression("solution") { ColumnSet = new ColumnSet("solutionid"), Criteria = new FilterExpression() };
                            defaultSolQuery.Criteria.AddCondition("uniquename", ConditionOperator.Equal, "Default");
                            var defaultSolResult = Service.RetrieveMultiple(defaultSolQuery);
                            if (defaultSolResult.Entities.Any()) defaultSolutionId = defaultSolResult.Entities.First().Id;
                        }
                        catch (Exception ex) { LogWarning("Failed to retrieve default solution ID: " + ex.Message); }

                        var scQ = new QueryExpression("solutioncomponent") { ColumnSet = new ColumnSet("solutionid"), Criteria = new FilterExpression() };
                        scQ.Criteria.AddCondition("componenttype", ConditionOperator.Equal, SolutionComponentTypeCloudFlow);
                        scQ.Criteria.AddCondition("objectid", ConditionOperator.Equal, flowId);
                        var solutionComponents = Service.RetrieveMultiple(scQ);

                        var solutionIds = solutionComponents.Entities.Select(sc => GetGuidFromAttribute(sc, "solutionid")).Where(g => g != Guid.Empty).Distinct().ToList();
                        var solutionNames = new Dictionary<Guid, string>();
                        if (solutionIds.Any())
                        {
                            var solQ = new QueryExpression("solution") { ColumnSet = new ColumnSet("solutionid", "friendlyname", "uniquename", "ismanaged") };
                            solQ.Criteria.AddCondition("solutionid", ConditionOperator.In, solutionIds.Cast<object>().ToArray());
                            var sols = Service.RetrieveMultiple(solQ);
                            foreach (var s in sols.Entities)
                            {
                                var id = s.Id; var friendly = s.GetAttributeValue<string>("friendlyname") ?? s.GetAttributeValue<string>("uniquename"); var uniq = s.GetAttributeValue<string>("uniquename") ?? string.Empty; if (!string.IsNullOrEmpty(uniq) && uniq.Equals("default", StringComparison.OrdinalIgnoreCase)) continue; if (!string.IsNullOrEmpty(friendly) && (friendly.IndexOf("default solution", StringComparison.OrdinalIgnoreCase) >= 0 || friendly.IndexOf("active solution", StringComparison.OrdinalIgnoreCase) >= 0)) continue; solutionNames[s.Id] = friendly;
                            }
                        }

                        if (solutionNames.Any()) fi.Solutions = string.Join(", ", solutionNames.Values.Distinct());
                        else fi.Solutions = null;

                        var solutionIdForLink = solutionIds.Any() ? solutionIds.First() : defaultSolutionId;
                        fi.LinkToFlow = solutionIdForLink != Guid.Empty && !string.IsNullOrEmpty(environmentId)
                            ? $"https://make.powerautomate.com/environments/{environmentId}/solutions/{solutionIdForLink}/flows/{flowId}?v3=true"
                            : string.Empty;
                        // --- End of fix ---

                        try
                        {
                            var req = new RetrieveSharedPrincipalsAndAccessRequest { Target = new EntityReference("workflow", flowId) };
                            var resp = (RetrieveSharedPrincipalsAndAccessResponse)Service.Execute(req);
                            var principalIds = new List<Guid>();
                            if (resp?.PrincipalAccesses != null) foreach (var pa in resp.PrincipalAccesses) if (pa.Principal is EntityReference er) principalIds.Add(er.Id);

                            // resolve principal names (users and teams) and detect disabled users
                            var principalNames = new Dictionary<Guid, string>();
                            var principalDisabled = new HashSet<Guid>();
                            if (principalIds.Any() || fi.OwnerId != Guid.Empty)
                            {
                                var idsToQuery = new HashSet<Guid>(principalIds);
                                if (fi.OwnerId != Guid.Empty) idsToQuery.Add(fi.OwnerId);
                                var idsArray = idsToQuery.ToArray();
                                try
                                {
                                    var userQ = new QueryExpression("systemuser") { ColumnSet = new ColumnSet("systemuserid", "fullname", "isdisabled") };
                                    userQ.Criteria.AddCondition("systemuserid", ConditionOperator.In, idsArray.Cast<object>().ToArray());
                                    var users = Service.RetrieveMultiple(userQ);
                                    foreach (var u in users.Entities)
                                    {
                                        var fullname = GetStringSafe(u, "fullname");
                                        principalNames[u.Id] = fullname ?? u.Id.ToString();
                                        bool isDisabled = false;
                                        try { if (u.Contains("isdisabled")) isDisabled = u.GetAttributeValue<bool>("isdisabled"); } catch { }
                                        try { if (!string.IsNullOrEmpty(fullname) && DisabledUserExceptions.Contains(fullname.Trim())) isDisabled = false; } catch { }
                                        if (isDisabled) principalDisabled.Add(u.Id);
                                    }
                                }
                                catch (Exception ex) { LogWarning("Failed to resolve user names: " + ex.Message); }

                                try
                                {
                                    var teamQ = new QueryExpression("team") { ColumnSet = new ColumnSet("teamid", "name") };
                                    teamQ.Criteria.AddCondition("teamid", ConditionOperator.In, idsArray.Cast<object>().ToArray());
                                    var teams = Service.RetrieveMultiple(teamQ);
                                    foreach (var t in teams.Entities) principalNames[t.Id] = GetStringSafe(t, "name");
                                }
                                catch (Exception ex) { LogWarning("Failed to resolve team names: " + ex.Message); }

                                // build Co-Owner display names excluding primary owner
                                if (principalIds != null && principalIds.Any())
                                {
                                    var coIds = principalIds.Where(id => id != fi.OwnerId).Distinct().ToList();
                                    var coNames = new List<string>();
                                    foreach (var id in coIds)
                                    {
                                        var name = principalNames.ContainsKey(id) ? principalNames[id] : id.ToString();
                                        if (principalDisabled.Contains(id)) name += " (disabled)";
                                        coNames.Add(name);
                                    }
                                    fi.CoOwners = coNames.Any() ? string.Join(", ", coNames) : null;
                                }
                                else fi.CoOwners = null;

                                // prefer resolved name for primary owner and annotate if disabled
                                try
                                {
                                    if (fi.OwnerId != Guid.Empty && principalNames.ContainsKey(fi.OwnerId)) fi.Owner = principalNames[fi.OwnerId];
                                    if (fi.OwnerId != Guid.Empty && principalDisabled.Contains(fi.OwnerId)) fi.Owner = (fi.Owner ?? string.Empty) + " (disabled)";
                                }
                                catch { }
                            }
                        }
                        catch (Exception ex) { a.Result = ex; return; }
                    }
                    catch (Exception ex) { a.Result = ex; return; }
                    a.Result = new { Flow = fi, DisabledPrincipalIds = disabledSet.ToArray() };
                },
                PostWorkCallBack = (a) =>
                {
                    if (a.Error != null) { MessageBox.Show(a.Error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                    if (a.Result is Exception ex) { MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                    dynamic res = a.Result; var fi = (FlowInfo)res.Flow; var disabled = (Guid[])res.DisabledPrincipalIds;
                    var idx = lastResults.FindIndex(x => x.Id == fi.Id); if (idx >= 0) lastResults[idx] = fi;
                    if (!string.IsNullOrEmpty(fi.ClientDataJson)) _flowClientData[fi.Id] = fi.ClientDataJson;
                    foreach (DataGridViewRow row in dgvFlows.Rows)
                    {
                        if (row.Tag is Guid id && id == fi.Id)
                        {
                            row.Cells["Name"].Value = fi.Name;
                            row.Cells["Status"].Value = fi.Status ?? "";
                            row.Cells["Link to Flow"].Value = fi.LinkToFlow ?? "";
                            row.Cells["Description"].Value = fi.Description ?? "";
                            row.Cells["Solutions"].Value = fi.Solutions ?? "";
                            row.Cells["Primary Owner"].Value = fi.Owner ?? "";
                            row.Cells["Co-Owners"].Value = fi.CoOwners ?? "";
                            try { row.Cells["Co-Owners"].ToolTipText = fi.CoOwners ?? string.Empty; } catch { }
                            try
                            {
                                bool hasDisabled = false;
                                // detect disabled markers in display strings but ignore any entries that are in the exception list
                                if (!string.IsNullOrEmpty(fi.Owner) && fi.Owner.IndexOf("(disabled)", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    var ownerBase = RemoveDisabledMarker(fi.Owner);
                                    if (!DisabledUserExceptions.Contains(ownerBase)) hasDisabled = true;
                                }
                                if (!hasDisabled && !string.IsNullOrEmpty(fi.CoOwners) && fi.CoOwners.IndexOf("(disabled)", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    try
                                    {
                                        var parts = fi.CoOwners.Split(',');
                                        foreach (var p in parts)
                                        {
                                            if (p.IndexOf("(disabled)", StringComparison.OrdinalIgnoreCase) >= 0)
                                            {
                                                var baseName = RemoveDisabledMarker(p);
                                                if (!DisabledUserExceptions.Contains(baseName)) { hasDisabled = true; break; }
                                            }
                                        }
                                    }
                                    catch { }
                                }

                                if (hasDisabled)
                                {
                                    row.DefaultCellStyle.BackColor = Color.LightYellow;
                                    row.DefaultCellStyle.SelectionBackColor = Color.Gold;
                                    row.DefaultCellStyle.ForeColor = Color.DarkRed;
                                    try { row.DefaultCellStyle.Font = new Font(dgvFlows.Font, FontStyle.Bold); } catch { }
                                }
                                else
                                {
                                    row.DefaultCellStyle.BackColor = Color.Empty;
                                    row.DefaultCellStyle.SelectionBackColor = Color.Empty;
                                    row.DefaultCellStyle.ForeColor = Color.Empty;
                                    try { row.DefaultCellStyle.Font = dgvFlows.Font; } catch { }
                                }
                            }
                            catch { }
                            break;
                        }
                    }
                }
            });
        }

        private bool TryOpenConnectionDialog()
        {
            MessageBox.Show(
                "Please connect to an environment using the XrmToolBox connection manager before loading flows.",
                "Not connected",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return false;
        }

        private void toolStripMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void dgvFlows_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            if (dgvFlows.Columns[e.ColumnIndex].Name == "Link to Flow")
            {
                var url = dgvFlows.Rows[e.RowIndex].Cells[e.ColumnIndex].Value as string;
                if (!string.IsNullOrEmpty(url) && Uri.IsWellFormedUriString(url, UriKind.Absolute))
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = url, UseShellExecute = true });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed to open link: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private static List<Entity> RetrieveAllPages(IOrganizationService svc, QueryExpression qe)
        {
            qe.PageInfo = new PagingInfo { Count = 5000, PageNumber = 1, ReturnTotalRecordCount = false };
            var all = new List<Entity>();
            EntityCollection page;
            do
            {
                page = svc.RetrieveMultiple(qe);
                all.AddRange(page.Entities);
                qe.PageInfo.PageNumber++;
                qe.PageInfo.PagingCookie = page.PagingCookie;
            } while (page.MoreRecords);
            return all;
        }

        private static IEnumerable<List<T>> ToBatches<T>(IEnumerable<T> source, int maxBatchSize)
        {
            var batch = new List<T>(maxBatchSize);
            foreach (var item in source)
            {
                batch.Add(item);
                if (batch.Count == maxBatchSize) { yield return batch; batch = new List<T>(maxBatchSize); }
            }
            if (batch.Count > 0) yield return batch;
        }
    }
}