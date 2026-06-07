using PmAi.RdpRemoteLauncher;
using Xunit;

namespace PmAiRdpRemoteLauncher.Tests;

public class LauncherIniTests
{
    [Fact]
    public void Load_parsesDisconnectOnChildExit()
    {
        var path = Path.Combine(Path.GetTempPath(), "rap-" + Guid.NewGuid().ToString("N") + ".ini");
        File.WriteAllText(
            path,
            """
            起動プログラム番号=1
            終了時RDP切断=0
            1="C:\app.exe"
            """);

        try
        {
            var ini = LauncherIni.Load(path);
            Assert.False(ini.DisconnectOnChildExit);
            Assert.False(ini.ResolveDisconnectOnChildExit());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("1", true)]
    [InlineData("0", false)]
    [InlineData("true", true)]
    [InlineData("off", false)]
    public void ResolveDisconnectOnChildExit_envOverridesIni(string envValue, bool expected)
    {
        var path = Path.Combine(Path.GetTempPath(), "rap-" + Guid.NewGuid().ToString("N") + ".ini");
        File.WriteAllText(
            path,
            """
            起動プログラム番号=1
            終了時RDP切断=1
            1="C:\app.exe"
            """);

        Environment.SetEnvironmentVariable(LauncherIni.DisconnectOnChildExitEnvVar, envValue);
        try
        {
            var ini = LauncherIni.Load(path);
            Assert.Equal(expected, ini.ResolveDisconnectOnChildExit());
        }
        finally
        {
            Environment.SetEnvironmentVariable(LauncherIni.DisconnectOnChildExitEnvVar, null);
            File.Delete(path);
        }
    }
}
