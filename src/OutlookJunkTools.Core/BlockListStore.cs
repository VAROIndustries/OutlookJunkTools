using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace OutlookJunkTools.Core
{
    /// <summary>Reads and writes the block list as JSON.</summary>
    public sealed class BlockListStore
    {
        [DataContract]
        private sealed class Dto
        {
            [DataMember(Name = "kind", Order = 1)] public string Kind;
            [DataMember(Name = "value", Order = 2)] public string Value;
            [DataMember(Name = "addedUtc", Order = 3)] public string AddedUtc;
        }

        private readonly string _path;

        public BlockListStore(string filePath) { _path = filePath; }

        public string FilePath => _path;

        /// <summary>Loads the list. <paramref name="warning"/> is set only when a corrupt file was set aside.</summary>
        public BlockList Load(out string warning)
        {
            warning = null;
            var list = new BlockList();
            if (!File.Exists(_path)) return list;
            try
            {
                List<Dto> dtos;
                using (var fs = File.OpenRead(_path))
                    dtos = (List<Dto>)new DataContractJsonSerializer(typeof(List<Dto>)).ReadObject(fs);
                foreach (var d in dtos ?? new List<Dto>())
                {
                    if (d == null || string.IsNullOrWhiteSpace(d.Value)) continue;
                    var kind = string.Equals(d.Kind, "domain", StringComparison.OrdinalIgnoreCase)
                        ? BlockKind.Domain : BlockKind.Address;
                    DateTime added;
                    if (!DateTime.TryParse(d.AddedUtc, CultureInfo.InvariantCulture,
                            DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out added))
                        added = DateTime.UtcNow;
                    list.Add(kind, d.Value, added);
                }
                return list;
            }
            catch (Exception ex)
            {
                var aside = Path.Combine(Path.GetDirectoryName(_path),
                    "blocklist.corrupt-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".json");
                try { File.Move(_path, aside); } catch { /* leave it; Save will overwrite */ }
                warning = "The block list file could not be read (" + ex.GetType().Name +
                          "). It was set aside as " + Path.GetFileName(aside) + " and an empty list is in use.";
                return new BlockList();
            }
        }

        public void Save(BlockList list)
        {
            var dir = Path.GetDirectoryName(_path);
            Directory.CreateDirectory(dir);
            var dtos = new List<Dto>();
            foreach (var e in list.Entries)
                dtos.Add(new Dto
                {
                    Kind = e.Kind == BlockKind.Domain ? "domain" : "address",
                    Value = e.Value,
                    AddedUtc = e.AddedUtc.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture)
                });

            // Write beside the target, then swap, so a crash cannot leave a half-written list.
            var temp = _path + ".tmp";
            using (var fs = File.Create(temp))
                new DataContractJsonSerializer(typeof(List<Dto>)).WriteObject(fs, dtos);
            if (File.Exists(_path)) File.Replace(temp, _path, null);
            else File.Move(temp, _path);
        }
    }
}
