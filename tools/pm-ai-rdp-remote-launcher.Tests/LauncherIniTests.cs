using PmAi.RdpRemoteLauncher;
using Xunit;

namespace PmAiRdpRemoteLauncher.Tests;

public class LauncherIniTests
{
    [Fact]
    public void Load_parsesDisabledSlot()
    {
        var path = Path.Combine(Path.GetTempPath(), "rap-" + Guid.NewGuid().ToString("N") + ".ini");
        File.WriteAllText(
            path,
            """
            起動プログラム番号=0
            終了時RDP切断=1
            2="C:\app.exe"
            """);

        try
        {
            var ini = LauncherIni.Load(path);
            Assert.True(ini.IsLauncherDisabled);
            Assert.Null(ini.ResolveSelectedCommand());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void WriteSelectedSlot_updatesHeadLinePreservingSlots()
    {
        var path = Path.Combine(Path.GetTempPath(), "rap-" + Guid.NewGuid().ToString("N") + ".ini");
        File.WriteAllText(
            path,
            """
            起動プログラム番号=2
            終了時RDP切断=1
            2="C:\app.exe" arg
            """,
            System.Text.Encoding.UTF8);

        try
        {
            LauncherIni.WriteSelectedSlot(path, LauncherIni.DisabledSlot);

            var text = File.ReadAllText(path, System.Text.Encoding.UTF8);
            Assert.Contains("起動プログラム番号=0", text);
            Assert.Contains("2=\"C:\\app.exe\" arg", text);
            var reloaded = LauncherIni.Load(path);
            Assert.True(reloaded.IsLauncherDisabled);
            Assert.Equal("\"C:\\app.exe\" arg", reloaded.Slots[2]);
        }
        finally
        {
            File.Delete(path);
        }
    }

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
