using System;
using System.Collections.Generic;
using System.Linq;

namespace OutlookJunkTools.Core
{
    public enum BlockKind { Address, Domain }

    public sealed class BlockEntry
    {
        public BlockKind Kind;
        public string Value;
        public DateTime AddedUtc;
    }

    /// <summary>Addresses and domains whose mail should go to Junk.</summary>
    public sealed class BlockList
    {
        private readonly List<BlockEntry> _entries = new List<BlockEntry>();

        public IReadOnlyList<BlockEntry> Entries => _entries;

        public static string Normalize(string value)
        {
            if (value == null) return "";
            value = value.Trim().ToLowerInvariant();
            return value.StartsWith("@") ? value.Substring(1) : value;
        }

        /// <summary>Returns the lower-cased domain of an SMTP address, or null if there is none.</summary>
        public static string DomainOf(string smtpAddress)
        {
            if (string.IsNullOrWhiteSpace(smtpAddress)) return null;
            var s = smtpAddress.Trim().ToLowerInvariant();
            int at = s.LastIndexOf('@');
            if (at < 0 || at == s.Length - 1) return null;
            return s.Substring(at + 1);
        }

        public bool Add(BlockKind kind, string value) => Add(kind, value, DateTime.UtcNow);

        public bool Add(BlockKind kind, string value, DateTime addedUtc)
        {
            var v = Normalize(value);
            if (v.Length == 0) return false;
            if (_entries.Any(e => e.Kind == kind && e.Value == v)) return false;
            _entries.Add(new BlockEntry { Kind = kind, Value = v, AddedUtc = addedUtc });
            return true;
        }

        public bool Remove(BlockKind kind, string value)
        {
            var v = Normalize(value);
            return _entries.RemoveAll(e => e.Kind == kind && e.Value == v) > 0;
        }

        public bool IsBlocked(string smtpAddress)
        {
            var domain = DomainOf(smtpAddress);
            if (domain == null) return false;
            var address = smtpAddress.Trim().ToLowerInvariant();
            foreach (var e in _entries)
            {
                if (e.Kind == BlockKind.Address)
                {
                    if (e.Value == address) return true;
                }
                else if (domain == e.Value || domain.EndsWith("." + e.Value, StringComparison.Ordinal))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
