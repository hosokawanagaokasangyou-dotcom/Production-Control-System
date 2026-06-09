using PmAi.RdpRemoteLauncher;
using Xunit;

namespace PmAiRdpRemoteLauncher.Tests;

public class ProcessRunningCheckerTests
{
    [Fact]
    public void ExtractScenarioPathFragment_fromScenarioFlag()
    {
        var fragment =
            ProcessRunningChecker.ExtractScenarioPathFragment(
                "--scenario \"\\\\192.168.0.101\\share\\scenarios\\工程別生産計画問い合わせ.ardrpa\"");
        Assert.Equal("工程別生産計画問い合わせ.ardrpa", fragment);
    }

    [Fact]
    public void ExtractScenarioPathFragment_fromBarePath()
    {
        var fragment =
            ProcessRunningChecker.ExtractScenarioPathFragment(
                "\\\\server\\share\\a.ardrpa");
        Assert.Equal("a.ardrpa", fragment);
    }

    [Fact]
    public void ExtractScenarioPathFragment_emptyWhenNoScenario()
    {
        Assert.Null(ProcessRunningChecker.ExtractScenarioPathFragment("--eternal"));
        Assert.Null(ProcessRunningChecker.ExtractScenarioPathFragment(null));
    }
}
