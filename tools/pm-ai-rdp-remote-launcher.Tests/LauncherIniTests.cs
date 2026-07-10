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
            Assert.True(ini.IsSignOutOnly);
            Assert.Equal(LauncherIni.SignOutSlot, ini.SelectedSlot);
            Assert.Null(ini.ResolveSelectedCommand());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Load_parsesSignOutSlot99()
    {
        var path = Path.Combine(Path.GetTempPath(), "rap-" + Guid.NewGuid().ToString("N") + ".ini");
        File.WriteAllText(
            path,
            """
            起動プログラム番号=99
            終了時RDP切断=1
            2="C:\app.exe"
            """);

        try
        {
            var ini = LauncherIni.Load(path);
            Assert.True(ini.IsSignOutOnly);
            Assert.Equal(99, ini.SelectedSlot);
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
            LauncherIni.WriteSelectedSlot(path, LauncherIni.SignOutSlot);

            var text = File.ReadAllText(path, System.Text.Encoding.UTF8);
            Assert.Contains("起動プログラム番号=99", text);
            Assert.Contains("2=\"C:\\app.exe\" arg", text);
            var reloaded = LauncherIni.Load(path);
            Assert.True(reloaded.IsSignOutOnly);
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

    [Fact]
    public void Load_parsesSessionEndActionDisconnect()
    {
        var path = Path.Combine(Path.GetTempPath(), "rap-" + Guid.NewGuid().ToString("N") + ".ini");
        File.WriteAllText(
            path,
            """
            起動プログラム番号=1
            終了時RDP切断=1
            終了時セッション操作=切断
            1="C:\app.exe"
            """);

        try
        {
            var ini = LauncherIni.Load(path);
            Assert.Equal(SessionEndAction.Disconnect, ini.ResolveSessionEndAction());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Load_legacyDisconnectFlagDefaultsToSignOut()
    {
        var path = Path.Combine(Path.GetTempPath(), "rap-" + Guid.NewGuid().ToString("N") + ".ini");
        File.WriteAllText(
            path,
            """
            起動プログラム番号=1
            終了時RDP切断=1
            1="C:\app.exe"
            """);

        try
        {
            var ini = LauncherIni.Load(path);
            Assert.Equal(SessionEndAction.SignOut, ini.ResolveSessionEndAction());
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

    [Fact]
    public void BuildIniFileNameForOperator_appendsBasename()
    {
        Assert.Equal("細川_RPA設定.ini", LauncherIni.BuildIniFileNameForOperator("細川"));
        Assert.Equal(LauncherIni.DefaultIniFileName, LauncherIni.BuildIniFileNameForOperator(""));
        Assert.Equal("bad_name_RPA設定.ini", LauncherIni.BuildIniFileNameForOperator("bad/name"));
    }

    [Fact]
    public void ResolveIniPathInDeployLayout_usesExeDirForOperatorIni()
    {
        var exeDir = Path.Combine(Path.GetTempPath(), "rdp-launcher-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(exeDir);
        try
        {
            var resolved = LauncherIni.ResolveIniPathInDeployLayout(exeDir, "細川");
            Assert.Equal(
                Path.Combine(exeDir, "細川_RPA設定.ini"),
                resolved);
        }
        finally
        {
            Directory.Delete(exeDir, recursive: true);
        }
    }

    [Fact]
    public void ResolveIniPathInDeployLayout_defaultsToExeDirRpaIni()
    {
        var exeDir = Path.Combine(Path.GetTempPath(), "rdp-launcher-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(exeDir);
        try
        {
            var resolved = LauncherIni.ResolveIniPathInDeployLayout(exeDir, null);
            Assert.Equal(
                Path.Combine(exeDir, LauncherIni.DefaultIniFileName),
                resolved);
        }
        finally
        {
            Directory.Delete(exeDir, recursive: true);
        }
    }

    [Fact]
    public void TryParseOperatorArgument_readsFirstPositional()
    {
        Assert.Equal("細川", LauncherIni.TryParseOperatorArgument(["細川"]));
        Assert.Equal("細川", LauncherIni.TryParseOperatorArgument(["--ini", "C:\\x.ini", "細川"]));
        Assert.Null(LauncherIni.TryParseOperatorArgument(["--ini", "C:\\x.ini"]));
    }
}
