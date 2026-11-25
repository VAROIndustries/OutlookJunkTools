using OutlookJunkTools.Core;
using Xunit;

public class AuthResultsTests
{
    [Fact] public void Pass() => Assert.Equal(DkimVerdict.Pass,
        AuthResults.Dkim("Authentication-Results: mx.example; spf=pass;\r\n dkim=pass header.d=x.com\r\n"));

    [Theory] [InlineData("dkim=fail")] [InlineData("dkim=none")]
    public void Not_pass(string clause) => Assert.Equal(DkimVerdict.NotPass,
        AuthResults.Dkim("Authentication-Results: mx.example; " + clause + "\r\n"));

    [Fact] public void Unknown_when_absent()
    {
        Assert.Equal(DkimVerdict.Unknown, AuthResults.Dkim("From: a@x.com\r\n"));
        Assert.Equal(DkimVerdict.Unknown, AuthResults.Dkim(null));
        Assert.Equal(DkimVerdict.Unknown, AuthResults.Dkim("Authentication-Results: mx.example; spf=pass\r\n"));
    }

    [Fact] public void Any_passing_header_wins() => Assert.Equal(DkimVerdict.Pass,
        AuthResults.Dkim("Authentication-Results: a; dkim=fail\r\nAuthentication-Results: b; dkim=pass\r\n"));
}
