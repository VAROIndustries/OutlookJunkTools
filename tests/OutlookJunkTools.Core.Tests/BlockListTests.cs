using OutlookJunkTools.Core;
using Xunit;

public class BlockListTests
{
    [Fact] public void Address_match_ignores_case()
    {
        var l = new BlockList(); l.Add(BlockKind.Address, "A@X.com");
        Assert.True(l.IsBlocked("a@x.com"));
        Assert.True(l.IsBlocked(" A@x.COM "));
        Assert.False(l.IsBlocked("b@x.com"));
    }

    [Fact] public void Domain_blocks_subdomains_but_not_lookalikes()
    {
        var l = new BlockList(); l.Add(BlockKind.Domain, "x.com");
        Assert.True(l.IsBlocked("a@x.com"));
        Assert.True(l.IsBlocked("a@mail.x.com"));
        Assert.False(l.IsBlocked("a@notx.com"));
        Assert.False(l.IsBlocked("a@x.com.evil.org"));
    }

    [Theory]
    [InlineData(null)] [InlineData("")] [InlineData("no-at-sign")] [InlineData("trailing@")]
    public void Unusable_addresses_are_never_blocked(string address)
    {
        var l = new BlockList(); l.Add(BlockKind.Domain, "x.com");
        Assert.False(l.IsBlocked(address));
    }

    [Fact] public void Add_and_remove_report_whether_they_changed_anything()
    {
        var l = new BlockList();
        Assert.True(l.Add(BlockKind.Domain, "@X.com"));
        Assert.False(l.Add(BlockKind.Domain, "x.com"));
        Assert.Equal("x.com", l.Entries[0].Value);
        Assert.True(l.Remove(BlockKind.Domain, "X.COM"));
        Assert.False(l.Remove(BlockKind.Domain, "x.com"));
        Assert.False(l.Add(BlockKind.Address, "  "));
    }

    [Fact] public void Same_value_can_exist_as_address_and_domain()
    {
        var l = new BlockList();
        Assert.True(l.Add(BlockKind.Domain, "x.com"));
        Assert.True(l.Add(BlockKind.Address, "x.com"));
    }

    [Fact] public void DomainOf_returns_null_without_a_domain()
    {
        Assert.Equal("x.com", BlockList.DomainOf("A@X.com"));
        Assert.Null(BlockList.DomainOf("nope"));
        Assert.Null(BlockList.DomainOf(null));
    }
}
