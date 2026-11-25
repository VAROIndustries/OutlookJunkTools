using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkTools.AddIn
{
    internal static class SenderResolver
    {
        private const string PrSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
        private const string PrSentRepresentingSmtp = "http://schemas.microsoft.com/mapi/proptag/0x5D02001F";

        /// <summary>Returns the sender's SMTP address, or null when none can be found.</summary>
        public static string GetSmtpAddress(Outlook.MailItem mail)
        {
            try
            {
                if (!string.Equals(mail.SenderEmailType, "EX", StringComparison.OrdinalIgnoreCase))
                    return Clean(mail.SenderEmailAddress);

                Outlook.AddressEntry sender = mail.Sender;
                if (sender != null)
                {
                    try
                    {
                        var user = sender.GetExchangeUser();
                        if (user != null && Clean(user.PrimarySmtpAddress) != null) return Clean(user.PrimarySmtpAddress);
                    }
                    catch { /* fall through to MAPI properties */ }
                    try { return Clean(sender.PropertyAccessor.GetProperty(PrSmtpAddress) as string); }
                    catch { /* fall through */ }
                }
                try { return Clean(mail.PropertyAccessor.GetProperty(PrSentRepresentingSmtp) as string); }
                catch { return null; }
            }
            catch (Exception ex)
            {
                Log.Error("Resolving sender", ex);
                return null;
            }
        }

        private static string Clean(string s)
        {
            if (string.IsNullOrWhiteSpace(s) || s.IndexOf('@') <= 0) return null;
            return s.Trim().ToLowerInvariant();
        }
    }
}
