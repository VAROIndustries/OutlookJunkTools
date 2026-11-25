using System;
using System.IO;
using System.Linq;
using OutlookJunkTools.Core;
using Xunit;

public class BlockListStoreTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "ojt-tests-" + Guid.NewGuid().ToString("N"));
    private string FilePath => Path.Combine(_dir, "sub", "blocklist.json");
    public void Dispose() { try { Directory.Delete(_dir, true); } catch { } }

    [Fact] public void Missing_file_gives_empty_list_and_no_warning()
    {
        var list = new BlockListStore(FilePath).Load(out var warning);
        Assert.Empty(list.Entries);
        Assert.Null(warning);
    }

    [Fact] public void Save_creates_folder_and_round_trips()
    {
        var store = new BlockListStore(FilePath);
        var list = new BlockList();
        var when = new DateTime(2026, 11, 25, 15, 30, 0, DateTimeKind.Utc);
        list.Add(BlockKind.Address, "a@x.com", when);
        list.Add(BlockKind.Domain, "junk.biz", when);
        store.Save(list);
        store.Save(list); // second save exercises the replace path

        var loaded = store.Load(out var warning);
        Assert.Null(warning);
        Assert.Equal(2, loaded.Entries.Count);
        Assert.Equal(BlockKind.Address, loaded.Entries[0].Kind);
        Assert.Equal("a@x.com", loaded.Entries[0].Value);
        Assert.Equal(BlockKind.Domain, loaded.Entries[1].Kind);
        Assert.Equal(when, loaded.Entries[1].AddedUtc);
        Assert.False(File.Exists(FilePath + ".tmp"));
    }

    [Fact] public void Corrupt_file_is_set_aside_and_saving_still_works()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
        File.WriteAllText(FilePath, "{not json");
        var store = new BlockListStore(FilePath);

        var list = store.Load(out var warning);
        Assert.Empty(list.Entries);
        Assert.NotNull(warning);
        Assert.Single(Directory.GetFiles(Path.GetDirectoryName(FilePath), "blocklist.corrupt-*.json"));

        list.Add(BlockKind.Domain, "x.com");
        store.Save(list);
        Assert.Single(store.Load(out warning).Entries);
        Assert.Null(warning);
    }
}
