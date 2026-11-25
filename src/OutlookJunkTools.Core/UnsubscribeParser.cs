using System;
using System.Collections.Generic;
using System.Net;
using System.Text.RegularExpressions;

namespace OutlookJunkTools.Core
{
    public enum UnsubscribeMethod { OneClickPost, Mailto, HeaderLink, BodyLink }

    public sealed class UnsubscribeOption
    {
        public UnsubscribeMethod Method;
        /// <summary>The URL, or for Mailto a mailto: URI.</summary>
        public Uri Target;
        public string MailAddress;
        public string MailSubject;
        public string MailBody;
    }

    /// <summary>Finds the ways a message offers to unsubscribe, best first.</summary>
    public static class UnsubscribeParser
    {
        private static readonly Regex AngleUri = new Regex("<([^<>]+)>", RegexOptions.Compiled);
        private static readonly Regex Anchor = new Regex(
            "<a\\b[^>]*?href\\s*=\\s*(?:\"([^\"]*)\"|'([^']*)')[^>]*>(.*?)</a>",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline);
        private static readonly Regex Tag = new Regex("<[^>]+>", RegexOptions.Compiled);
        private static readonly string[] Keywords = { "unsubscribe", "opt out", "opt-out", "optout", "manage preferences" };

        public static IReadOnlyList<UnsubscribeOption> Parse(string rawHeaders, string htmlBody)
        {
            var options = new List<UnsubscribeOption>();

            var header = HeaderReader.Get(rawHeaders, "List-Unsubscribe");
            var post = HeaderReader.Get(rawHeaders, "List-Unsubscribe-Post");
            bool oneClick = post != null &&
                string.Equals(post.Replace(" ", ""), "List-Unsubscribe=One-Click", StringComparison.OrdinalIgnoreCase);

            Uri https = null, web = null;
            UnsubscribeOption mailto = null;
            if (header != null)
            {
                foreach (Match m in AngleUri.Matches(header))
                {
                    // Folding can leave whitespace inside a URI.
                    var text = Regex.Replace(m.Groups[1].Value, "\\s+", "");
                    Uri uri;
                    if (!Uri.TryCreate(text, UriKind.Absolute, out uri)) continue;
                    if (uri.Scheme == Uri.UriSchemeHttps) { if (https == null) https = uri; if (web == null) web = uri; }
                    else if (uri.Scheme == Uri.UriSchemeHttp) { if (web == null) web = uri; }
                    else if (uri.Scheme == Uri.UriSchemeMailto) { if (mailto == null) mailto = ParseMailto(uri, text); }
                    // Any other scheme is ignored.
                }
            }

            if (oneClick && https != null)
                options.Add(new UnsubscribeOption { Method = UnsubscribeMethod.OneClickPost, Target = https });
            if (mailto != null) options.Add(mailto);
            if (web != null && !(oneClick && https != null && web == https))
                options.Add(new UnsubscribeOption { Method = UnsubscribeMethod.HeaderLink, Target = web });

            var body = FindBodyLink(htmlBody);
            if (body != null) options.Add(new UnsubscribeOption { Method = UnsubscribeMethod.BodyLink, Target = body });
            return options;
        }

        private static UnsubscribeOption ParseMailto(Uri uri, string text)
        {
            var rest = text.Substring("mailto:".Length);
            string query = "";
            int q = rest.IndexOf('?');
            if (q >= 0) { query = rest.Substring(q + 1); rest = rest.Substring(0, q); }
            var address = Uri.UnescapeDataString(rest).Trim();
            if (address.IndexOf('@') <= 0) return null;

            var option = new UnsubscribeOption
            {
                Method = UnsubscribeMethod.Mailto,
                Target = uri,
                MailAddress = address,
                MailSubject = "unsubscribe",
                MailBody = ""
            };
            foreach (var pair in query.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int eq = pair.IndexOf('=');
                if (eq <= 0) continue;
                var key = pair.Substring(0, eq);
                var value = Uri.UnescapeDataString(pair.Substring(eq + 1));
                if (key.Equals("subject", StringComparison.OrdinalIgnoreCase)) option.MailSubject = value;
                else if (key.Equals("body", StringComparison.OrdinalIgnoreCase)) option.MailBody = value;
            }
            return option;
        }

        private static Uri FindBodyLink(string html)
        {
            if (string.IsNullOrEmpty(html)) return null;
            foreach (Match m in Anchor.Matches(html))
            {
                var href = WebUtility.HtmlDecode(m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value).Trim();
                var text = WebUtility.HtmlDecode(Tag.Replace(m.Groups[3].Value, " "));
                if (!ContainsKeyword(text) && !ContainsKeyword(href)) continue;
                Uri uri;
                if (!Uri.TryCreate(href, UriKind.Absolute, out uri)) continue;
                if (uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeHttp) return uri;
            }
            return null;
        }

        private static bool ContainsKeyword(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            foreach (var k in Keywords)
                if (s.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return false;
        }
    }
}
