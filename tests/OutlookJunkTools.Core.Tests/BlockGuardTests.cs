using OutlookJunkTools.Core;
using Xunit;

public class BlockGuardTests
{
    private static BlockGuard Guard() => new BlockGuard(new[] { "Me@MyCo.com", "me@gmail.com" });

    [Fact] public void Own_address_is_refused() =>
        Assert.Equal(GuardResult.RefusedOwnAddress, Guard().Check(BlockKind.Address, "me@myco.com"));

    [Fact] public void Own_domain_needs_confirmation() =>
        Assert.Equal(GuardResult.NeedsConfirmation, Guard().Check(BlockKind.Domain, "myco.com"));

    [Theory]
    [InlineData("gmail.com")] [InlineData("outlook.com")] [InlineData("hotmail.com")] [InlineData("yahoo.com")]
    [InlineData("icloud.com")] [InlineData("aol.com")] [InlineData("proton.me")]
    public void Shared_providers_need_confirmation(string domain) =>
        Assert.Equal(GuardResult.NeedsConfirmation, new BlockGuard(new string[0]).Check(BlockKind.Domain, domain));

    [Fact] public void Everything_else_is_allowed()
    {
        Assert.Equal(GuardResult.Allowed, Guard().Check(BlockKind.Address, "spam@junk.biz"));
        Assert.Equal(GuardResult.Allowed, Guard().Check(BlockKind.Domain, "junk.biz"));
        Assert.Equal(GuardResult.Allowed, Guard().Check(BlockKind.Address, "someone@gmail.com"));
    }
}
