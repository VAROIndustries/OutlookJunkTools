using System;
using System.Collections.Generic;
using System.IO;
using OutlookJunkTools.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkTools.AddIn
{
    /// <summary>Everything the add-in shares between its parts. Used on Outlook's main thread only.</summary>
    internal sealed class AddInState
    {
        public Outlook.Application App;
        public BlockListStore Store;
        public BlockList List;
        public BlockGuard Guard;
        public string LoadWarning;

        public static AddInState Create(Outlook.Application app)
        {
            var state = new AddInState { App = app };
            state.Store = new BlockListStore(Path.Combine(Log.Directory, "blocklist.json"));
            state.List = state.Store.Load(out state.LoadWarning);
            state.Guard = new BlockGuard(OwnAddresses(app));
            return state;
        }

        public void Save() => Store.Save(List);

        private static List<string> OwnAddresses(Outlook.Application app)
        {
            var result = new List<string>();
            try
            {
                foreach (Outlook.Account account in app.Session.Accounts)
                {
                    try { if (!string.IsNullOrEmpty(account.SmtpAddress)) result.Add(account.SmtpAddress); }
                    catch (Exception ex) { Log.Error("Reading account address", ex); }
                }
            }
            catch (Exception ex) { Log.Error("Enumerating accounts", ex); }
            return result;
        }
    }
}
