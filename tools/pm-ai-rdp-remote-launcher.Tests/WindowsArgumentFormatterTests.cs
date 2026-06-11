using Xunit;

namespace PmAi.RdpRemoteLauncher.Tests;

public class WindowsArgumentFormatterTests
{
    [Fact]
    public void TokenizeForProcess_singleUncPathWithSpaces()
    {
        var arg = @"\\192.168.0.101\share\scenarios\工程別生産計画問い合わ せ.ardrpa";
        var tokens = WindowsArgumentFormatter.TokenizeForProcess(arg);
        Assert.Single(tokens);
        Assert.Equal(arg, tokens[0]);
    }

    [Fact]
    public void TokenizeForProcess_quotedUncPathWithSpaces()
    {
        var arg = @"""\\192.168.0.101\share\scenarios\工程別生産計画問い合わ せ.ardrpa""";
        var tokens = WindowsArgumentFormatter.TokenizeForProcess(arg);
        Assert.Single(tokens);
        Assert.Equal(@"\\192.168.0.101\share\scenarios\工程別生産計画問い合わ せ.ardrpa", tokens[0]);
    }

    [Fact]
    public void FormatArgumentString_quotesTokenWithSpaces()
    {
        var arg = @"\\192.168.0.101\share\scenarios\工程別生産計画問い合わ せ.ardrpa";
        var formatted = WindowsArgumentFormatter.FormatArgumentString(arg);
        Assert.Equal(@"""\\192.168.0.101\share\scenarios\工程別生産計画問い合わ せ.ardrpa""", formatted);
    }

    [Fact]
    public void FormatArgumentTokens_doesNotResplitScenarioPathSpaces()
    {
        var path = @"\\192.168.0.101\share\002  加工G\scenarios\a.ardrpa";
        var formatted = WindowsArgumentFormatter.FormatArgumentTokens(
            new[] { AladdinRpaLaunchArgs.ScenarioFlag, path });
        Assert.Contains("002  加工G", formatted);
        var tokens = WindowsArgumentFormatter.TokenizeForProcess(formatted);
        Assert.Equal(2, tokens.Count);
        Assert.Equal(path, tokens[1]);
    }

    [Fact]
    public void FormatArgumentString_leavesSimplePathUnquoted()
    {
        var arg = @"C:\Users\test\file.ardrpa";
        Assert.Equal(arg, WindowsArgumentFormatter.FormatArgumentString(arg));
    }

    [Fact]
    public void FormatArgumentString_preservesExistingQuotes()
    {
        var arg = @"""\\server\share\file.ardrpa""";
        Assert.Equal(arg, WindowsArgumentFormatter.FormatArgumentString(arg));
    }
}
