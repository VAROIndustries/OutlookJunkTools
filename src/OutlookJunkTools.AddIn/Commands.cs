using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using OutlookJunkTools.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkTools.AddIn
{
    /// <summary>What the ribbon buttons do.</summary>
    internal sealed class Commands
    {
        private const string Title = "Outlook Junk Tools";
        private readonly AddInState _state;

        public Commands(AddInState state) { _state = state; }

        /// <summary>The mail items a ribbon or context-menu click applies to.</summary>
        public List<Outlook.MailItem> Targets(object control)
        {
            var result = new List<Outlook.MailItem>();
            object context = null;
            try { context = ((dynamic)control).Context; } catch { }

            var selection = context as Outlook.Selection;
            var inspector = context as Outlook.Inspector;
            if (selection == null && inspector == null)
            {
                var explorer = context as Outlook.Explorer ?? _state.App.ActiveExplorer();
                if (explorer != null) { try { selection = explorer.Selection; } catch { } }
            }

            if (inspector != null)
            {
                var mail = inspector.CurrentItem as Outlook.MailItem;
                if (mail != null) result.Add(mail);
            }
            else if (selection != null)
            {
                foreach (object o in selection)
                {
                    var mail = o as Outlook.MailItem;
                    if (mail != null) result.Add(mail);
                }
            }
            return result;
        }

        public void Block(object control, BlockKind kind, bool alsoJunk)
        {
            var mails = Targets(control);
            if (mails.Count == 0) { Info("Select one or more email messages first."); return; }

            int added = 0, already = 0, own = 0, unresolved = 0, declined = 0, junked = 0, junkFailed = 0;
            var decided = new Dictionary<string, bool>();

            foreach (var mail in mails)
            {
                var address = SenderResolver.GetSmtpAddress(mail);
                if (address == null) { unresolved++; continue; }
                var value = kind == BlockKind.Domain ? BlockList.DomainOf(address) : address;

                bool proceed;
                if (!decided.TryGetValue(value, out proceed))
                {
                    proceed = true;
                    // Never block one of the user's own addresses, by address or as the only reason for a domain block.
                    if (_state.Guard.Check(BlockKind.Address, address) == GuardResult.RefusedOwnAddress)
                    {
                        own++; proceed = false;
                    }
                    else if (_state.Guard.Check(kind, value) == GuardResult.NeedsConfirmation)
                    {
                        proceed = MessageBox.Show(
                            "Block every sender at " + value + "?\n\n" +
                            "This is a shared mail provider or one of your own domains, so a lot of mail you want may come from it.",
                            Title, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes;
                        if (!proceed) declined++;
                    }
                    if (proceed)
                    {
                        if (_state.List.Add(kind, value)) { added++; Log.Info("Blocked " + kind + " " + value); }
                        else already++;
                    }
                    decided[value] = proceed;
                }

                if (proceed && alsoJunk)
                {
                    string error;
                    if (JunkMover.MoveToJunk(mail, out error)) junked++; else junkFailed++;
                }
            }

            if (added > 0) _state.Save();

            var noun = kind == BlockKind.Domain ? "domain" : "sender";
            var parts = new List<string>();
            parts.Add("Blocked " + added + " " + noun + (added == 1 ? "" : "s"));
            if (already > 0) parts.Add(already + " already blocked");
            if (own > 0) parts.Add(own + " skipped (your own address)");
            if (declined > 0) parts.Add(declined + " skipped (you chose not to)");
            if (unresolved > 0) parts.Add(unresolved + " skipped (no sender address found)");
            if (alsoJunk) parts.Add(junked + " moved to Junk");
            if (junkFailed > 0) parts.Add(junkFailed + " could not be moved (see log)");
            var message = string.Join(", ", parts) + ".";
            if (decided.Count == 1 && added == 1) message += "\n\n" + decided.Keys.First();
            Info(message);
        }

        public void Junk(object control)
        {
            var mails = Targets(control);
            if (mails.Count == 0) { Info("Select one or more email messages first."); return; }
            var errors = new List<string>();
            foreach (var mail in mails)
            {
                string error;
                if (!JunkMover.MoveToJunk(mail, out error)) errors.Add(error);
            }
            if (errors.Count > 0)
                Info(errors.Count + " message(s) could not be moved to Junk.\n\n" + string.Join("\n", errors.Distinct()));
        }

        public void Unsubscribe(object control)
        {
            var mails = Targets(control);
            if (mails.Count != 1) { Info("Select a single email message to unsubscribe from."); return; }
            UnsubscribeRunner.Run(_state, mails[0]);
        }

        public void ManageList()
        {
            using (var form = new ManageListForm(_state)) form.ShowDialog();
        }

        public static void Info(string message) =>
            MessageBox.Show(message, Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
