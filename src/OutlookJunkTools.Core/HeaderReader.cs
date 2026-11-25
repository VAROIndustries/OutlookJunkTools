using System;
using System.Collections.Generic;
using System.Text;

namespace OutlookJunkTools.Core
{
    /// <summary>Reads values out of a raw internet header block, unfolding continuation lines.</summary>
    public static class HeaderReader
    {
        public static string Get(string rawHeaders, string name)
        {
            var all = GetAll(rawHeaders, name);
            return all.Count > 0 ? all[0] : null;
        }

        public static IReadOnlyList<string> GetAll(string rawHeaders, string name)
        {
            var result = new List<string>();
            if (string.IsNullOrEmpty(rawHeaders) || string.IsNullOrEmpty(name)) return result;

            var lines = rawHeaders.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var prefix = name + ":";
            StringBuilder current = null;
            foreach (var line in lines)
            {
                bool continuation = line.Length > 0 && (line[0] == ' ' || line[0] == '\t');
                if (continuation)
                {
                    if (current != null) current.Append(' ').Append(line.Trim());
                    continue;
                }
                if (current != null) { result.Add(current.ToString().Trim()); current = null; }
                if (line.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    current = new StringBuilder(line.Substring(prefix.Length).Trim());
            }
            if (current != null) result.Add(current.ToString().Trim());
            return result;
        }
    }
}
