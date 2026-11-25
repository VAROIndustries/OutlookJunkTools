using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using OutlookJunkTools.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookJunkTools.AddIn
{
    internal static class UnsubscribeRunner
    {
        private const string PrTransportHeaders = "http://schemas.microsoft.com/mapi/proptag/0x007D001F";

        public static void Run(AddInState state, Outlook.MailItem mail)
        {
            string headers = null, html = null;
            try { headers = mail.PropertyAccessor.GetProperty(PrTransportHeaders) as string; }
            catch (Exception ex) { Log.Error("Reading headers", ex); }
            try { html = mail.HTMLBody; } catch (Exception ex) { Log.Error("Reading body", ex); }

            var options = UnsubscribeParser.Parse(headers, html);
            if (options.Count == 0)
            {
                Commands.Info("This message has no unsubscribe header and no unsubscribe link that could be found.");
                return;
            }

            var sender = SenderResolver.GetSmtpAddress(mail) ?? "(unknown sender)";
            using (var form = new ConfirmUnsubscribeForm(state, mail, sender, options, AuthResults.Dkim(headers)))
                form.ShowDialog();
        }

        /// <summary>RFC 8058 one-click request. Runs off the UI thread. Returns null on success, else an error message.</summary>
        public static string PostOneClick(Uri target)
        {
            if (target == null || target.Scheme != Uri.UriSchemeHttps) return "Refused: one-click unsubscribe must use https.";
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(target);
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                request.AllowAutoRedirect = false;
                request.Timeout = 15000;
                request.ReadWriteTimeout = 15000;
                request.CookieContainer = null;
                request.UseDefaultCredentials = false;
                request.UserAgent = "OutlookJunkTools/0.2";
                var body = Encoding.ASCII.GetBytes("List-Unsubscribe=One-Click");
                request.ContentLength = body.Length;
                using (var stream = request.GetRequestStream()) stream.Write(body, 0, body.Length);
                using (var response = (HttpWebResponse)request.GetResponse())
                {
                    int code = (int)response.StatusCode;
                    return code >= 200 && code < 300 ? null : "The server answered " + code + " " + response.StatusDescription + ".";
                }
            }
            catch (WebException ex)
            {
                var response = ex.Response as HttpWebResponse;
                if (response != null)
                    return "The server answered " + (int)response.StatusCode + " " + response.StatusDescription + ".";
                return ex.Message;
            }
            catch (Exception ex) { return ex.Message; }
        }

        /// <summary>Sends the unsubscribe mail from the account that received the message. Returns null on success.</summary>
        public static string SendMailto(AddInState state, Outlook.MailItem original, UnsubscribeOption option)
        {
            try
            {
                var reply = (Outlook.MailItem)state.App.CreateItem(Outlook.OlItemType.olMailItem);
                reply.To = option.MailAddress;
                reply.Subject = string.IsNullOrEmpty(option.MailSubject) ? "unsubscribe" : option.MailSubject;
                reply.BodyFormat = Outlook.OlBodyFormat.olFormatPlain;
                reply.Body = option.MailBody ?? "";

                var account = AccountFor(state, original);
                if (account != null) reply.SendUsingAccount = account;
                reply.Send();
                return null;
            }
            catch (Exception ex)
            {
                Log.Error("SendMailto", ex);
                return ex.Message;
            }
        }

        public static string OpenInBrowser(Uri target)
        {
            if (target == null || (target.Scheme != Uri.UriSchemeHttps && target.Scheme != Uri.UriSchemeHttp))
                return "Refused: only http and https links are opened.";
            try
            {
                Process.Start(new ProcessStartInfo(target.AbsoluteUri) { UseShellExecute = true });
                return null;
            }
            catch (Exception ex) { return ex.Message; }
        }

        private static Outlook.Account AccountFor(AddInState state, Outlook.MailItem mail)
        {
            try
            {
                var storeId = ((Outlook.Folder)mail.Parent).Store.StoreID;
                foreach (Outlook.Account account in state.App.Session.Accounts)
                {
                    try { if (account.DeliveryStore != null && account.DeliveryStore.StoreID == storeId) return account; }
                    catch { }
                }
            }
            catch (Exception ex) { Log.Error("Finding account", ex); }
            return null;
        }
    }
}
