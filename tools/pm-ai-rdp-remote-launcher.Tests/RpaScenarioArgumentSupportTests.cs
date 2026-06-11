using PmAi.RdpRemoteLauncher;
using Xunit;

namespace PmAiRdpRemoteLauncher.Tests;

public class RpaScenarioArgumentSupportTests
{
    [Fact]
    public void RepairScenarioArguments_preservesDoubleSpaceInUnquotedPath()
    {
        var path = @"\\192.168.0.101\share\002  加工G\a.ardrpa";
        var repaired =
            RpaScenarioArgumentSupport.RepairScenarioArguments("--scenario " + path);
        Assert.Equal(
            "--scenario \"" + path + "\"",
            repaired);
    }

    [Fact]
    public void RepairScenarioArguments_fixesCollapsed002Segment()
    {
        var wrong = @"\\192.168.0.101\share\002 加工G\a.ardrpa";
        var correct = @"\\192.168.0.101\share\002  加工G\a.ardrpa";
        var repaired =
            RpaScenarioArgumentSupport.RepairScenarioArguments("--scenario " + wrong);
        Assert.Equal(
            "--scenario \"" + correct + "\"",
            repaired);
    }

    [Fact]
    public void RepairScenarioArguments_roundTripPreserves002DoubleSpace()
    {
        var wrong = @"\\192.168.0.101\share\002 加工G\a.ardrpa";
        var correct = @"\\192.168.0.101\share\002  加工G\a.ardrpa";
        var repaired =
            RpaScenarioArgumentSupport.RepairScenarioArguments("--scenario " + wrong);
        Assert.Contains("002  加工G", repaired);
        var paths = RpaScenarioArgumentSupport.ExtractScenarioPaths(repaired);
        Assert.Single(paths);
        Assert.Equal(correct, paths[0]);
    }

    [Fact]
    public void RepairScenarioArguments_preservesEternalFlag()
    {
        var path = @"\\192.168.0.101\share\002 加工G\a.ardrpa";
        var repaired =
            RpaScenarioArgumentSupport.RepairScenarioArguments(
                "--scenario " + path + " --eternal");
        Assert.Contains("--eternal", repaired);
        Assert.Contains("002  加工G", repaired);
    }
}
