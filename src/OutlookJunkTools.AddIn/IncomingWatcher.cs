using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkTools.AddIn
{
    /// <summary>Moves newly arrived mail from blocked senders to Junk.</summary>
    internal sealed class IncomingWatcher
    {
        private const int SweepDays = 7;
        private const int SweepMaxPerInbox = 300;

        private readonly AddInState _state;
        // Held in fields so the COM event sources are not garbage collected.
        private readonly List<Outlook.Items> _inboxItems = new List<Outlook.Items>();
        private Timer _sweepTimer;

        public IncomingWatcher(AddInState state) { _state = state; }

        public void Start()
        {
            _state.App.NewMailEx += OnNewMailEx;
            foreach (var inbox in Inboxes())
            {
                try
                {
                    var items = inbox.Items;
                    items.ItemAdd += OnItemAdd;
                    _inboxItems.Add(items);
                }
                catch (Exception ex) { Log.Error("Hooking inbox", ex); }
            }
            Log.Info("Watching " + _inboxItems.Count + " inbox(es)");

            // Sweep shortly after startup so Outlook's own startup is not slowed down.
            _sweepTimer = new Timer { Interval = 20000 };
            _sweepTimer.Tick += (s, e) =>
            {
                _sweepTimer.Stop();
                try { Sweep(); } catch (Exception ex) { Log.Error("Startup sweep", ex); }
            };
            _sweepTimer.Start();
        }

        public void Stop()
        {
            try { _sweepTimer?.Stop(); } catch { }
            try { _state.App.NewMailEx -= OnNewMailEx; } catch { }
            foreach (var items in _inboxItems) { try { items.ItemAdd -= OnItemAdd; } catch { } }
            _inboxItems.Clear();
        }

        private IEnumerable<Outlook.MAPIFolder> Inboxes()
        {
            var result = new List<Outlook.MAPIFolder>();
            try
            {
                foreach (Outlook.Store store in _state.App.Session.Stores)
                {
                    try { result.Add(store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)); }
                    catch { /* store has no inbox (archive, public folders) */ }
                }
            }
            catch (Exception ex) { Log.Error("Enumerating stores", ex); }
            return result;
        }

        private void OnNewMailEx(string entryIdCollection)
        {
            try
            {
                if (string.IsNullOrEmpty(entryIdCollection)) return;
                foreach (var id in entryIdCollection.Split(','))
                {
                    try { Check(_state.App.Session.GetItemFromID(id.Trim()) as Outlook.MailItem, "new mail"); }
                    catch (Exception ex) { Log.Error("NewMailEx item", ex); }
                }
            }
            catch (Exception ex) { Log.Error("NewMailEx", ex); }
        }

        private void OnItemAdd(object item)
        {
            try { Check(item as Outlook.MailItem, "inbox add"); }
            catch (Exception ex) { Log.Error("ItemAdd", ex); }
        }

        private void Sweep()
        {
            if (_state.List.Entries.Count == 0) return;
            var cutoff = DateTime.Now.AddDays(-SweepDays);
            int moved = 0;
            foreach (var inbox in Inboxes())
            {
                try
                {
                    var items = inbox.Items.Restrict("[UnRead] = true");
                    items.Sort("[ReceivedTime]", true);
                    var candidates = new List<Outlook.MailItem>();
                    int seen = 0;
                    foreach (object o in items)
                    {
                        if (++seen > SweepMaxPerInbox) break;
                        var mail = o as Outlook.MailItem;
                        if (mail == null) continue;
                        if (mail.ReceivedTime < cutoff) break;
                        candidates.Add(mail);
                    }
                    // Move after enumerating; moving while enumerating skips items.
                    foreach (var mail in candidates) if (Check(mail, "startup sweep")) moved++;
                }
                catch (Exception ex) { Log.Error("Sweeping an inbox", ex); }
            }
            Log.Info("Startup sweep moved " + moved + " message(s)");
        }

        private bool Check(Outlook.MailItem mail, string source)
        {
            if (mail == null || _state.List.Entries.Count == 0) return false;
            var address = SenderResolver.GetSmtpAddress(mail);
            if (address == null || !_state.List.IsBlocked(address)) return false;
            string error;
            bool ok = JunkMover.MoveToJunk(mail, out error);
            Log.Info("Blocked sender " + address + " (" + source + "): " + (ok ? "moved to Junk" : "not moved - " + error));
            return ok;
        }
    }
}
