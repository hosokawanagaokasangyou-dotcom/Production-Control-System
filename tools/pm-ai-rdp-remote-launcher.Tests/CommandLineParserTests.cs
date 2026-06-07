using System.Text;
using Xunit;

namespace PmAi.RdpRemoteLauncher.Tests;

public class CommandLineParserTests
{
    [Fact]
    public void Parse_simpleExe()
    {
        var parsed = CommandLineParser.Parse(@"C:\Windows\System32\notepad.exe");
        Assert.Equal(@"C:\Windows\System32\notepad.exe", parsed.Executable);
        Assert.Equal(string.Empty, parsed.Arguments);
    }

    [Fact]
    public void Parse_uncWithArgs()
    {
        var line = @"\\server\share\Aladdin_RPA_Studio.exe C:\Users\test\file.ardrpa";
        var parsed = CommandLineParser.Parse(line);
        Assert.Equal(@"\\server\share\Aladdin_RPA_Studio.exe", parsed.Executable);
        Assert.Equal(@"C:\Users\test\file.ardrpa", parsed.Arguments);
    }

    [Fact]
    public void Parse_quotedExecutable()
    {
        var line = @"""C:\Program Files\App\app.exe"" --flag";
        var parsed = CommandLineParser.Parse(line);
        Assert.Equal(@"C:\Program Files\App\app.exe", parsed.Executable);
        Assert.Equal("--flag", parsed.Arguments);
    }
}

public class LauncherIniTests
{
    [Fact]
    public void Load_readsSelectedSlot()
    {
        var path = Path.GetTempFileName();
        try
        {
            File.WriteAllText(
                path,
                """
                起動プログラム番号=2
                1=C:\Windows\System32\notepad.exe
                2=\\server\share\app.exe arg
                """,
                Encoding.UTF8);
            var ini = LauncherIni.Load(path);
            Assert.Equal(2, ini.SelectedSlot);
            Assert.Equal(@"\\server\share\app.exe arg", ini.ResolveSelectedCommand());
        }
        finally
        {
            File.Delete(path);
        }
    }
}
