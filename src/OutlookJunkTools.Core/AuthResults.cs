using System.Text.RegularExpressions;

namespace OutlookJunkTools.Core
{
    public enum DkimVerdict { Pass, NotPass, Unknown }

    /// <summary>
    /// Reports what the receiving mail server recorded about DKIM. This repeats the server's
    /// verdict from the Authentication-Results header; it does not verify the signature.
    /// </summary>
    public static class AuthResults
    {
        private static readonly Regex Pass = new Regex("\\bdkim\\s*=\\s*pass\\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex Any = new Regex("\\bdkim\\s*=", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static DkimVerdict Dkim(string rawHeaders)
        {
            bool sawDkim = false;
            foreach (var h in HeaderReader.GetAll(rawHeaders, "Authentication-Results"))
            {
                if (Pass.IsMatch(h)) return DkimVerdict.Pass;
                if (Any.IsMatch(h)) sawDkim = true;
            }
            return sawDkim ? DkimVerdict.NotPass : DkimVerdict.Unknown;
        }
    }
}
