using PmAi.RdpRemoteLauncher;
using Xunit;

namespace PmAiRdpRemoteLauncher.Tests;

public class AladdinRpaArgumentAppenderTests
{
    [Fact]
    public void AppendCredentials_putsCredentialsBeforeEternal()
    {
        var tokens = AladdinRpaArgumentAppender.AppendCredentials(
            new[] { AladdinRpaLaunchArgs.EternalFlag },
            new OperatorAladdinCredentials("000585", "secret"));

        Assert.Equal(
            new[]
            {
                AladdinRpaLaunchArgs.IdFlag,
                "000585",
                AladdinRpaLaunchArgs.PasswordFlag,
                "secret",
                AladdinRpaLaunchArgs.EternalFlag,
            },
            tokens);
    }

    [Fact]
    public void AppendCredentials_putsScenarioBetweenCredentialsAndEternal()
    {
        var tokens = AladdinRpaArgumentAppender.AppendCredentials(
            new[]
            {
                AladdinRpaLaunchArgs.ScenarioFlag,
                @"\\server\share\a.ardrpa",
                AladdinRpaLaunchArgs.EternalFlag,
            },
            new OperatorAladdinCredentials("000585", "secret"));

        Assert.Equal(
            new[]
            {
                AladdinRpaLaunchArgs.IdFlag,
                "000585",
                AladdinRpaLaunchArgs.PasswordFlag,
                "secret",
                AladdinRpaLaunchArgs.ScenarioFlag,
                @"\\server\share\a.ardrpa",
                AladdinRpaLaunchArgs.EternalFlag,
            },
            tokens);
    }

    [Fact]
    public void AppendCredentials_normalizesLegacyBareScenarioPath()
    {
        var tokens = AladdinRpaArgumentAppender.AppendCredentials(
            new[] { @"\\server\share\a.ardrpa" },
            new OperatorAladdinCredentials("000585", "secret"));

        Assert.Equal(
            new[]
            {
                AladdinRpaLaunchArgs.IdFlag,
                "000585",
                AladdinRpaLaunchArgs.PasswordFlag,
                "secret",
                AladdinRpaLaunchArgs.ScenarioFlag,
                @"\\server\share\a.ardrpa",
            },
            tokens);
    }

    [Fact]
    public void AppendCredentials_withoutScenario_usesCredentialsOnly()
    {
        var tokens = AladdinRpaArgumentAppender.AppendCredentials(
            Array.Empty<string>(),
            new OperatorAladdinCredentials("000585", "secret"));

        Assert.Equal(
            new[]
            {
                AladdinRpaLaunchArgs.IdFlag,
                "000585",
                AladdinRpaLaunchArgs.PasswordFlag,
                "secret",
            },
            tokens);
    }
}
