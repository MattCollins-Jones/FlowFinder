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
    internal sealed class ManageCoOwnersDialog : Form
    {
        private IOrganizationService _service; private Guid _flowId; private Action<string> _logInfo; private Action<string> _logWarn;
        private ListBox lbCoOwners; private ComboBox cbUsers; private Button btnAdd; private Button btnRemove; private Button btnClose;
        private Task _lastOperation = Task.CompletedTask;
        public Task LastOperation => _lastOperation;
        public bool ChangesMade { get; private set; } = false;

        public ManageCoOwnersDialog(IOrganizationService service, Guid flowId, Action<string> logInfo = null, Action<string> logWarn = null)
        {
            _service = service; _flowId = flowId; _logInfo = logInfo ?? (_ => { }); _logWarn = logWarn ?? (_ => { }); Initialize(); LoadData();
        }

        private void Initialize()
        {
            this.Text = "Manage Co-Owners"; this.Width = 600; this.Height = 400; this.StartPosition = FormStartPosition.CenterParent;
            lbCoOwners = new ListBox() { Left = 10, Top = 10, Width = 350, Height = 300 }; this.Controls.Add(lbCoOwners);
            cbUsers = new ComboBox() { Left = 370, Top = 10, Width = 200, DropDownStyle = ComboBoxStyle.DropDown, AutoCompleteMode = AutoCompleteMode.SuggestAppend, AutoCompleteSource = AutoCompleteSource.ListItems }; this.Controls.Add(cbUsers);
            btnAdd = new Button() { Left = 370, Top = 50, Width = 200, Text = "Add as Co-Owner" }; btnAdd.Click += BtnAdd_Click; this.Controls.Add(btnAdd);
            btnRemove = new Button() { Left = 370, Top = 90, Width = 200, Text = "Remove Selected Co-Owner" }; btnRemove.Click += BtnRemove_Click; this.Controls.Add(btnRemove);
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
            catch (Exception ex) { _logWarn("Failed to load Co-Owners: " + ex.Message); MessageBox.Show("Failed to load Co-Owners: " + ex.Message); }

            var names = new Dictionary<Guid, string>();
            var disabledSetLocal = new HashSet<Guid>();
            if (principalIds.Any())
            {
                try
                {
                    var userQ = new QueryExpression("systemuser") { ColumnSet = new ColumnSet("systemuserid", "fullname", "isdisabled") };
                    userQ.Criteria.AddCondition("systemuserid", ConditionOperator.In, principalIds.Cast<object>().ToArray());
                    var users = _service.RetrieveMultiple(userQ);
                    foreach (var u in users.Entities)
                    {
                        var fullname = u.GetAttributeValue<string>("fullname");
                        names[u.Id] = fullname ?? u.Id.ToString();
                        bool isDisabled = false;
                        try { if (u.Contains("isdisabled")) isDisabled = u.GetAttributeValue<bool>("isdisabled"); } catch { }
                        // Do not treat listed exception names as disabled (trim whitespace)
                        try { if (!string.IsNullOrEmpty(fullname) && FlowFinderControl.DisabledUserExceptions.Contains(fullname.Trim())) isDisabled = false; } catch { }
                        if (isDisabled) disabledSetLocal.Add(u.Id);
                    }
                }
                catch (Exception ex) { _logWarn("Failed to resolve Co-Owner user names: " + ex.Message); }
                try { var teamQ = new QueryExpression("team") { ColumnSet = new ColumnSet("teamid", "name") }; teamQ.Criteria.AddCondition("teamid", ConditionOperator.In, principalIds.Cast<object>().ToArray()); var teams = _service.RetrieveMultiple(teamQ); foreach (var t in teams.Entities) names[t.Id] = t.GetAttributeValue<string>("name"); }
                catch (Exception ex) { _logWarn("Failed to resolve team names: " + ex.Message); }
            }

            foreach (var id in principalIds)
            {
                var display = names.ContainsKey(id) ? names[id] : id.ToString();
                if (disabledSetLocal.Contains(id)) display += " (disabled)";
                lbCoOwners.Items.Add(new ListItem { Id = id, Name = display });
            }

            // Populate users combobox with available users — use paging to handle environments with >5000 users
            try
            {
                var available = new List<ListItem>();
                var userQAll = new QueryExpression("systemuser") { ColumnSet = new ColumnSet("systemuserid", "fullname") };
                userQAll.Criteria.AddCondition("isdisabled", ConditionOperator.Equal, false);
                userQAll.PageInfo = new PagingInfo { Count = 5000, PageNumber = 1, ReturnTotalRecordCount = false };
                EntityCollection page;
                do
                {
                    page = _service.RetrieveMultiple(userQAll);
                    foreach (var u in page.Entities)
                    {
                        var uid = u.Id;
                        if (principalIds.Contains(uid)) continue;
                        var fullname = string.Empty;
                        try { fullname = u.GetAttributeValue<string>("fullname") ?? uid.ToString(); } catch { fullname = uid.ToString(); }
                        available.Add(new ListItem { Id = uid, Name = fullname });
                    }
                    userQAll.PageInfo.PageNumber++;
                    userQAll.PageInfo.PagingCookie = page.PagingCookie;
                } while (page.MoreRecords);

                available = available.OrderBy(x => x.Name).ToList();
                if (available.Any())
                {
                    cbUsers.DisplayMember = "Name";
                    cbUsers.ValueMember = "Id";
                    cbUsers.DataSource = available;
                }
            }
            catch (Exception ex)
            {
                _logWarn("Failed to load available users: " + ex.Message);
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            var sel = cbUsers.SelectedItem as ListItem; if (sel == null) { MessageBox.Show("Select a user to add."); return; }
            btnAdd.Enabled = false; btnRemove.Enabled = false; cbUsers.Enabled = false; lbCoOwners.Enabled = false; btnClose.Enabled = false;
            var busy = new BusyForm("Adding Co-Owner..."); busy.Show(this);
            var op = Task.Run(() =>
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
                try { busy.Close(); } catch { /* safe cleanup */ }
                if (IsDisposed || !IsHandleCreated) return;
                btnAdd.Enabled = true; btnRemove.Enabled = true; cbUsers.Enabled = true; lbCoOwners.Enabled = true; btnClose.Enabled = true;
                if (t.Result == null) { _logInfo($"Added Co-Owner {sel.Name}"); MessageBox.Show("Added Co-Owner."); ChangesMade = true; }
                else { _logWarn("Failed to add Co-Owner: " + t.Result.Message); MessageBox.Show("Failed to add Co-Owner: " + t.Result.Message); }
                if (!IsDisposed) LoadData();
            }, TaskScheduler.FromCurrentSynchronizationContext());
            _lastOperation = op;
        }

        private void BtnRemove_Click(object sender, EventArgs e)
        {
            var sel = lbCoOwners.SelectedItem as ListItem; if (sel == null) { MessageBox.Show("Select a Co-Owner to remove."); return; }
            var confirm = MessageBox.Show($"Are you sure you want to remove Co-Owner '{sel.Name}'?", "Confirm remove", MessageBoxButtons.YesNo, MessageBoxIcon.Question); if (confirm != DialogResult.Yes) return;
            btnAdd.Enabled = false; btnRemove.Enabled = false; cbUsers.Enabled = false; lbCoOwners.Enabled = false; btnClose.Enabled = false;
            var busy = new BusyForm("Removing Co-Owner..."); busy.Show(this);
            var op = Task.Run(() =>
            {
                try { var revoke = new RevokeAccessRequest { Target = new EntityReference("workflow", _flowId), Revokee = new EntityReference("systemuser", sel.Id) }; _service.Execute(revoke); return (Exception)null; }
                catch (Exception ex) { return ex; }
            }).ContinueWith(t =>
            {
                try { busy.Close(); } catch { /* safe cleanup */ }
                if (IsDisposed || !IsHandleCreated) return;
                btnAdd.Enabled = true; btnRemove.Enabled = true; cbUsers.Enabled = true; lbCoOwners.Enabled = true; btnClose.Enabled = true;
                if (t.Result == null) { _logInfo($"Removed Co-Owner {sel.Name}"); MessageBox.Show("Removed Co-Owner."); ChangesMade = true; }
                else { _logWarn("Failed to remove Co-Owner: " + t.Result.Message); MessageBox.Show("Failed to remove Co-Owner: " + t.Result.Message); }
                if (!IsDisposed) LoadData();
            }, TaskScheduler.FromCurrentSynchronizationContext());
            _lastOperation = op;
        }

        private class ListItem { public Guid Id { get; set; } public string Name { get; set; } public override string ToString() => Name; }
    }
}
