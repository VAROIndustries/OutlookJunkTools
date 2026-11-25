using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OutlookJunkTools.Core;

namespace OutlookJunkTools.AddIn
{
    internal sealed class ManageListForm : Form
    {
        private readonly AddInState _state;
        private readonly ListView _view = new ListView();
        private readonly TextBox _input = new TextBox();

        public ManageListForm(AddInState state)
        {
            _state = state;
            Text = "Outlook Junk Tools - Block List";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(560, 420);
            MinimizeBox = false; ShowInTaskbar = false;
            Font = SystemFonts.MessageBoxFont;

            _view.SetBounds(12, 12, 536, 320);
            _view.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _view.View = View.Details; _view.FullRowSelect = true; _view.HideSelection = false;
            _view.Columns.Add("Blocked", 290); _view.Columns.Add("Type", 80); _view.Columns.Add("Added", 150);

            var remove = new Button { Text = "Remove selected", Left = 12, Top = 342, Width = 130, Height = 28,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            remove.Click += (s, e) => RemoveSelected();

            var hint = new Label { Text = "Add an address (name@example.com) or a domain (example.com):", Left = 12, Top = 380,
                Width = 536, Anchor = AnchorStyles.Bottom | AnchorStyles.Left };
            hint.Top = 376;
            _input.SetBounds(12, 394, 330, 24); _input.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            var add = new Button { Text = "Add", Left = 350, Top = 392, Width = 70, Height = 26, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            add.Click += (s, e) => AddEntry();
            var close = new Button { Text = "Close", Left = 470, Top = 392, Width = 78, Height = 26, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            close.Click += (s, e) => Close();
            ClientSize = new Size(560, 430);
            AcceptButton = add; CancelButton = close;

            Controls.AddRange(new Control[] { _view, remove, hint, _input, add, close });
            Reload();
        }

        private void Reload()
        {
            _view.BeginUpdate();
            _view.Items.Clear();
            foreach (var e in _state.List.Entries.OrderBy(x => x.Value))
            {
                var item = new ListViewItem(e.Value) { Tag = e };
                item.SubItems.Add(e.Kind == BlockKind.Domain ? "Domain" : "Address");
                item.SubItems.Add(e.AddedUtc.ToLocalTime().ToString("g"));
                _view.Items.Add(item);
            }
            _view.EndUpdate();
        }

        private void RemoveSelected()
        {
            try
            {
                if (_view.SelectedItems.Count == 0) return;
                foreach (ListViewItem item in _view.SelectedItems)
                {
                    var e = (BlockEntry)item.Tag;
                    _state.List.Remove(e.Kind, e.Value);
                    Log.Info("Unblocked " + e.Kind + " " + e.Value);
                }
                _state.Save();
                Reload();
            }
            catch (Exception ex) { Log.Error("Remove", ex); MessageBox.Show(this, ex.Message, Text); }
        }

        private void AddEntry()
        {
            try
            {
                var value = BlockList.Normalize(_input.Text);
                if (value.Length == 0 || value.IndexOf('.') < 0 || value.Contains(" "))
                {
                    MessageBox.Show(this, "Enter an email address or a domain.", Text); return;
                }
                var kind = value.Contains("@") ? BlockKind.Address : BlockKind.Domain;
                var verdict = _state.Guard.Check(kind, value);
                if (verdict == GuardResult.RefusedOwnAddress)
                {
                    MessageBox.Show(this, "That is one of your own addresses.", Text); return;
                }
                if (verdict == GuardResult.NeedsConfirmation && MessageBox.Show(this,
                        "Block every sender at " + value + "? A lot of mail you want may come from it.", Text,
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
                    return;
                if (!_state.List.Add(kind, value)) { MessageBox.Show(this, "Already on the list.", Text); return; }
                _state.Save();
                Log.Info("Blocked " + kind + " " + value + " (manual)");
                _input.Clear();
                Reload();
            }
            catch (Exception ex) { Log.Error("Add", ex); MessageBox.Show(this, ex.Message, Text); }
        }
    }
}
