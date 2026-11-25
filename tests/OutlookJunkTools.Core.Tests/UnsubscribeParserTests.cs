using System.Linq;
using OutlookJunkTools.Core;
using Xunit;

public class UnsubscribeParserTests
{
    private const string Both = "List-Unsubscribe: <mailto:u@x.com?subject=unsub>, <https://x.com/u?id=1>\r\n";
    private const string Post = "List-Unsubscribe-Post: List-Unsubscribe=One-Click\r\n";

    private static UnsubscribeMethod[] Methods(string headers, string body = null) =>
        UnsubscribeParser.Parse(headers, body).Select(o => o.Method).ToArray();

    [Fact] public void One_click_comes_first_when_the_post_header_is_present()
    {
        var options = UnsubscribeParser.Parse("From: a@x.com\r\n" + Both + Post, "");
        Assert.Equal(new[] { UnsubscribeMethod.OneClickPost, UnsubscribeMethod.Mailto }, options.Select(o => o.Method));
        Assert.Equal("https://x.com/u?id=1", options[0].Target.AbsoluteUri);
        Assert.Equal("u@x.com", options[1].MailAddress);
        Assert.Equal("unsub", options[1].MailSubject);
    }

    [Fact] public void Without_the_post_header_the_link_is_browser_only() =>
        Assert.Equal(new[] { UnsubscribeMethod.Mailto, UnsubscribeMethod.HeaderLink }, Methods(Both));

    [Fact] public void Plain_http_is_never_one_click() =>
        Assert.Equal(new[] { UnsubscribeMethod.HeaderLink },
            Methods("List-Unsubscribe: <http://x.com/u>\r\n" + Post));

    [Fact] public void Folded_header_parses_like_an_unfolded_one()
    {
        var folded = "List-Unsubscribe: <mailto:u@x.com?subject=unsub>,\r\n\t<https://x.com/u\r\n ?id=1>\r\n" + Post;
        var options = UnsubscribeParser.Parse(folded, null);
        Assert.Equal(new[] { UnsubscribeMethod.OneClickPost, UnsubscribeMethod.Mailto }, options.Select(o => o.Method));
        Assert.Equal("https://x.com/u?id=1", options[0].Target.AbsoluteUri);
    }

    [Fact] public void Hostile_schemes_are_ignored() =>
        Assert.Empty(Methods("List-Unsubscribe: <javascript:alert(1)>, <file:///c:/x>, <ftp://x.com/u>\r\n" + Post));

    [Theory]
    [InlineData("<a href=\"https://x.com/opt\">Unsubscribe</a>")]
    [InlineData("<A HREF='https://x.com/opt'><span>click to OPT OUT</span></A>")]
    [InlineData("<a style=\"x\" href=\"https://x.com/opt\">Manage Preferences</a>")]
    [InlineData("<a href=\"https://x.com/unsubscribe?a=1&amp;b=2\">here</a>")]
    public void Body_link_is_found(string html) =>
        Assert.Equal(new[] { UnsubscribeMethod.BodyLink }, Methods("From: a@x.com\r\n", "<p>hi</p>" + html));

    [Fact] public void Body_link_decodes_html_entities()
    {
        var o = UnsubscribeParser.Parse(null, "<a href=\"https://x.com/unsubscribe?a=1&amp;b=2\">here</a>");
        Assert.Equal("https://x.com/unsubscribe?a=1&b=2", o[0].Target.AbsoluteUri);
    }

    [Fact] public void Body_link_with_a_bad_scheme_is_ignored() =>
        Assert.Empty(Methods(null, "<a href=\"javascript:void(0)\">Unsubscribe</a>"));

    [Fact] public void Nothing_found_gives_an_empty_list()
    {
        Assert.Empty(Methods("From: a@x.com\r\n", "<a href=\"https://x.com/shop\">Shop now</a>"));
        Assert.Empty(Methods(null, null));
    }

    [Fact] public void Mailto_subject_and_body_are_decoded()
    {
        var o = UnsubscribeParser.Parse("List-Unsubscribe: <mailto:u@x.com?subject=a%20b&body=c%0Ad>\r\n", null);
        Assert.Equal("a b", o[0].MailSubject);
        Assert.Equal("c\nd", o[0].MailBody);
    }

    [Fact] public void Header_name_match_ignores_case_and_similar_names()
    {
        var headers = "X-List-Unsubscribe: <https://bad.example/x>\r\nlist-unsubscribe: <https://x.com/u>\r\n";
        var o = UnsubscribeParser.Parse(headers, null);
        Assert.Single(o);
        Assert.Equal("https://x.com/u", o[0].Target.AbsoluteUri);
    }
}
