using PmAi.RdpRemoteLauncher;
using Xunit;

namespace PmAiRdpRemoteLauncher.Tests;

public class ProgramSignOutTests
{
    [Fact]
    public void Load_suppressSlot_doesNotResolveCommand()
    {
        var path = Path.Combine(Path.GetTempPath(), "rap-" + Guid.NewGuid().ToString("N") + ".ini");
        File.WriteAllText(
            path,
            """
            起動プログラム番号=0
            操作者=細川
            1="C:\app.exe"
            """);

        try
        {
            var ini = LauncherIni.Load(path);
            Assert.True(ini.IsSuppressOnly);
            Assert.Null(ini.ResolveSelectedCommand());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Load_signOutSlot99_resolvesSignOutCommand()
    {
        var path = Path.Combine(Path.GetTempPath(), "rap-" + Guid.NewGuid().ToString("N") + ".ini");
        File.WriteAllText(
            path,
            """
            起動プログラム番号=99
            99=--signout
            操作者=細川
            """);

        try
        {
            var ini = LauncherIni.Load(path);
            Assert.True(ini.IsSignOutSlotSelected);
            Assert.Equal("--signout", ini.ResolveSelectedCommand());
            Assert.True(LauncherIni.IsSignOutSlotCommand(ini.ResolveSelectedCommand()));
        }
        finally
        {
            File.Delete(path);
        }
    }
}
