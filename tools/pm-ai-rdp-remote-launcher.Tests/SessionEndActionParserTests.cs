using PmAi.RdpRemoteLauncher;
using Xunit;

namespace PmAiRdpRemoteLauncher.Tests;

public class SessionEndActionParserTests
{
    [Fact]
    public void Parse_recognizesSignOut()
    {
        Assert.Equal(SessionEndAction.SignOut, SessionEndActionParser.Parse("サインアウト", SessionEndAction.None));
        Assert.Equal(SessionEndAction.SignOut, SessionEndActionParser.Parse("signout", SessionEndAction.None));
    }

    [Fact]
    public void Parse_recognizesDisconnect()
    {
        Assert.Equal(SessionEndAction.Disconnect, SessionEndActionParser.Parse("切断", SessionEndAction.None));
        Assert.Equal(SessionEndAction.Disconnect, SessionEndActionParser.Parse("disconnect", SessionEndAction.None));
    }

    [Fact]
    public void Parse_recognizesNone()
    {
        Assert.Equal(SessionEndAction.None, SessionEndActionParser.Parse("なし", SessionEndAction.SignOut));
        Assert.Equal(SessionEndAction.None, SessionEndActionParser.Parse("none", SessionEndAction.SignOut));
    }

    [Fact]
    public void Parse_unknown_returnsDefault()
    {
        Assert.Equal(
            SessionEndAction.Disconnect,
            SessionEndActionParser.Parse("unknown", SessionEndAction.Disconnect));
    }
}
