using System;
using System.Collections.Generic;

namespace OutlookJunkTools.Core
{
    public enum GuardResult { Allowed, RefusedOwnAddress, NeedsConfirmation }

    /// <summary>Stops the user blocking themselves, and flags domain blocks that would catch a lot of legitimate mail.</summary>
    public sealed class BlockGuard
    {
        private static readonly string[] SharedProviders =
        {
            "gmail.com", "googlemail.com", "outlook.com", "hotmail.com", "live.com", "msn.com",
            "yahoo.com", "icloud.com", "me.com", "aol.com", "proton.me", "protonmail.com"
        };

        private readonly HashSet<string> _ownAddresses = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _sensitiveDomains = new HashSet<string>(SharedProviders, StringComparer.Ordinal);

        public BlockGuard(IEnumerable<string> ownAddresses)
        {
            foreach (var a in ownAddresses ?? new string[0])
            {
                var n = BlockList.Normalize(a);
                if (n.Length == 0) continue;
                _ownAddresses.Add(n);
                var d = BlockList.DomainOf(n);
                if (d != null) _sensitiveDomains.Add(d);
            }
        }

        public GuardResult Check(BlockKind kind, string value)
        {
            var v = BlockList.Normalize(value);
            if (kind == BlockKind.Address)
                return _ownAddresses.Contains(v) ? GuardResult.RefusedOwnAddress : GuardResult.Allowed;
            return _sensitiveDomains.Contains(v) ? GuardResult.NeedsConfirmation : GuardResult.Allowed;
        }
    }
}
