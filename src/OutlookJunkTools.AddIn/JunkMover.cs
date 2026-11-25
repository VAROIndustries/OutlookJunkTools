using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkTools.AddIn
{
    internal static class JunkMover
    {
        /// <summary>Moves a message to the Junk folder of the account it lives in.</summary>
        public static bool MoveToJunk(Outlook.MailItem mail, out string error)
        {
            error = null;
            try
            {
                var parent = mail.Parent as Outlook.Folder;
                if (parent == null) { error = "The message is not in a mail folder."; return false; }

                Outlook.MAPIFolder junk;
                try { junk = parent.Store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk); }
                catch { junk = null; }
                if (junk == null) { error = "This account has no Junk folder."; return false; }
                if (junk.EntryID == parent.EntryID) return true; // already there

                mail.Move(junk);
                return true;
            }
            catch (Exception ex)
            {
                Log.Error("MoveToJunk", ex);
                error = ex.Message;
                return false;
            }
        }
    }
}
