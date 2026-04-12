using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Flow_Finder
{
    internal sealed class ManageSolutionsDialog : Form
    {
        private IOrganizationService _service; private Guid _flowId; private Action<string> _logInfo; private Action<string> _logWarn;
        private ListBox lbSolutions; private ComboBox cbAvailableSolutions; private Button btnAddToSolution; private Button btnRemoveFromSolution; private Button btnClose;
        private Task _lastOperation = Task.CompletedTask;
        public Task LastOperation => _lastOperation;
        public bool ChangesMade { get; private set; } = false;

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
                scQ.Criteria.AddCondition("componenttype", ConditionOperator.Equal, FlowFinderControl.SolutionComponentTypeCloudFlow);
                scQ.Criteria.AddCondition("objectid", ConditionOperator.Equal, _flowId);
                var scRes = _service.RetrieveMultiple(scQ);

                Guid LocalGetGuidFromAttribute(Entity e, string attr) { if (e == null || string.IsNullOrEmpty(attr) || !e.Attributes.ContainsKey(attr)) return Guid.Empty; var val = e.Attributes[attr]; if (val == null) return Guid.Empty; if (val is Guid) return (Guid)val; if (val is EntityReference) return ((EntityReference)val).Id; if (val is AliasedValue av) { if (av.Value is Guid) return (Guid)av.Value; if (av.Value is EntityReference) return ((EntityReference)av.Value).Id; } return Guid.Empty; }

                var solutionIds = scRes.Entities.Select(sc => LocalGetGuidFromAttribute(sc, "solutionid")).Where(g => g != Guid.Empty).Distinct().ToList();
                var currentSolutions = new List<Tuple<Guid, string, string, bool>>();
                if (solutionIds.Any())
                {
                    var solQ = new QueryExpression("solution") { ColumnSet = new ColumnSet("solutionid", "friendlyname", "uniquename", "ismanaged") };
                    solQ.Criteria.AddCondition("solutionid", ConditionOperator.In, solutionIds.Cast<object>().ToArray());
                    var sols = _service.RetrieveMultiple(solQ);
                    foreach (var s in sols.Entities)
                    {
                        var id = s.Id; var friendly = s.GetAttributeValue<string>("friendlyname") ?? s.GetAttributeValue<string>("uniquename"); var uniq = s.GetAttributeValue<string>("uniquename") ?? string.Empty; if (!string.IsNullOrEmpty(uniq) && uniq.Equals("default", StringComparison.OrdinalIgnoreCase)) continue; if (!string.IsNullOrEmpty(friendly) && (friendly.IndexOf("default solution", StringComparison.OrdinalIgnoreCase) >= 0 || friendly.IndexOf("active solution", StringComparison.OrdinalIgnoreCase) >= 0)) continue; currentSolutions.Add(Tuple.Create(id, friendly, uniq, s.GetAttributeValue<bool?>("ismanaged") ?? false));
                    }
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
            var op = Task.Run(() => { try { var req = new AddSolutionComponentRequest { SolutionUniqueName = sel.UniqueName, ComponentType = FlowFinderControl.SolutionComponentTypeCloudFlow, ComponentId = _flowId }; _service.Execute(req); return (Exception)null; } catch (Exception ex) { return ex; } }).ContinueWith(t => { try { busy.Close(); } catch { /* safe cleanup */ } if (t.Result == null) { MessageBox.Show("Flow added to solution."); ChangesMade = true; } else { MessageBox.Show("Failed to add flow to solution: " + t.Result.Message); } LoadData(); }, TaskScheduler.FromCurrentSynchronizationContext());
            _lastOperation = op;
        }

        private void BtnRemoveFromSolution_Click(object sender, EventArgs e)
        {
            var sel = lbSolutions.SelectedItem as SolutionItem; if (sel == null) { MessageBox.Show("Select a solution to remove the flow from."); return; }
            var confirm = MessageBox.Show($"Remove flow from solution '{sel.FriendlyName}'?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question); if (confirm != DialogResult.Yes) return;
            var busy = new BusyForm("Removing flow from solution..."); busy.Show(this);
            var op = Task.Run(() => { try { var req = new RemoveSolutionComponentRequest { SolutionUniqueName = sel.UniqueName, ComponentType = FlowFinderControl.SolutionComponentTypeCloudFlow, ComponentId = _flowId }; _service.Execute(req); return (Exception)null; } catch (Exception ex) { return ex; } }).ContinueWith(t => { try { busy.Close(); } catch { /* safe cleanup */ } if (t.Result == null) { MessageBox.Show("Flow removed from solution."); ChangesMade = true; } else { MessageBox.Show("Failed to remove flow from solution: " + t.Result.Message); } LoadData(); }, TaskScheduler.FromCurrentSynchronizationContext());
            _lastOperation = op;
        }

        private class SolutionItem { public Guid Id { get; set; } public string FriendlyName { get; set; } public string UniqueName { get; set; } public bool IsManaged { get; set; } public override string ToString() => FriendlyName; }
    }
}
