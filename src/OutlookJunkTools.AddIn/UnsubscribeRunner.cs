using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Security.Authentication;
using System.Threading.Tasks;
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
                return Post(target, SslProtocols.Tls12 | SslProtocols.Tls13);
            }
            catch (Exception ex) when (ex is NotSupportedException || ex is ArgumentException)
            {
                // TLS 1.3 is not available on this Windows build.
                try { return Post(target, SslProtocols.Tls12); }
                catch (Exception inner) { return Describe(inner); }
            }
            catch (Exception ex) { return Describe(ex); }
        }

        // Outlook hosts add-ins with .NET's legacy TLS defaults (SSL 3 / TLS 1.0), which modern servers refuse.
        // The protocol is set on this one handler so no process-wide setting is changed for other add-ins.
        private static string Post(Uri target, SslProtocols protocols)
        {
            using (var handler = new HttpClientHandler
            {
                SslProtocols = protocols,
                AllowAutoRedirect = false,
                UseCookies = false,
                UseDefaultCredentials = false
            })
            using (var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(15) })
            using (var content = new StringContent("List-Unsubscribe=One-Click", Encoding.ASCII, "application/x-www-form-urlencoded"))
            {
                client.DefaultRequestHeaders.UserAgent.ParseAdd("OutlookJunkTools/0.2");
                using (var response = client.PostAsync(target, content).GetAwaiter().GetResult())
                {
                    int code = (int)response.StatusCode;
                    return code >= 200 && code < 300 ? null : "The server answered " + code + " " + response.ReasonPhrase + ".";
                }
            }
        }

        private static string Describe(Exception ex)
        {
            if (ex is TaskCanceledException) return "The server did not answer within 15 seconds.";
            while (ex.InnerException != null) ex = ex.InnerException;
            return ex.Message;
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
