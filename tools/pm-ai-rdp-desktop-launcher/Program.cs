using System.Diagnostics;

namespace PmAi.RdpDesktopLauncher;

/// <summary>
/// Unicode-safe launcher for the portable JavaFX bundle. jpackage native exe breaks
/// <c>$APPDIR</c> on non-ASCII install paths; this stub resolves the exe folder and spawns
/// <c>runtime\bin\java.exe</c> like <c>launch-pm-ai-rdp-launcher.bat</c>.
/// </summary>
internal static class Program
{
    private const string MainClass = "jp.co.pm.ai.desktop.RemoteDesktopFxApp";

    public static int Main(string[] args)
    {
        var root = LauncherPaths.ResolveExecutableDirectory();
        if (string.IsNullOrWhiteSpace(root))
        {
            WriteError("Cannot resolve launcher install folder.");
            return 1;
        }

        var javaExe = Path.Combine(root, "runtime", "bin", "java.exe");
        if (!File.Exists(javaExe))
        {
            var javaHome = Environment.GetEnvironmentVariable("JAVA_HOME");
            if (!string.IsNullOrWhiteSpace(javaHome))
            {
                var alt = Path.Combine(javaHome.Trim(), "bin", "java.exe");
                if (File.Exists(alt))
                {
                    javaExe = alt;
                }
            }
        }

        if (!File.Exists(javaExe))
        {
            WriteError($"Java not found: {Path.Combine(root, "runtime", "bin", "java.exe")}");
            return 1;
        }

        var appDir = Path.Combine(root, "app");
        if (!Directory.Exists(appDir))
        {
            WriteError($"Missing app folder: {appDir}");
            return 1;
        }

        var jfxMod = Path.Combine(appDir, "jfx-mod");
        var classpath = Path.Combine(appDir, "*");

        var javaArgs = new List<string>
        {
            "-Dfile.encoding=UTF-8",
            "-Xms512m",
            "-Xmx2g",
            "-Dprism.order=sw",
            "--add-opens=javafx.base/com.sun.javafx.event=ALL-UNNAMED",
            "--add-opens=javafx.controls/javafx.scene.control.skin=ALL-UNNAMED",
            "--add-opens=javafx.controls/com.sun.javafx.scene.control.behavior=ALL-UNNAMED",
            "--add-exports=javafx.controls/com.sun.javafx.scene.control.behavior=ALL-UNNAMED",
            "--enable-native-access=javafx.graphics",
            "--module-path",
            jfxMod,
            "--add-modules",
            "javafx.controls,javafx.fxml,javafx.graphics,javafx.base,javafx.media,javafx.swing",
            "-classpath",
            classpath,
            MainClass,
        };
        foreach (var arg in args)
        {
            if (!string.IsNullOrEmpty(arg))
            {
                javaArgs.Add(arg);
            }
        }

        var psi = new ProcessStartInfo
        {
            FileName = javaExe,
            WorkingDirectory = root,
            UseShellExecute = false,
        };
        foreach (var opt in javaArgs)
        {
            psi.ArgumentList.Add(opt);
        }

        try
        {
            using var proc = Process.Start(psi);
            if (proc is null)
            {
                WriteError("Failed to start Java process.");
                return 1;
            }
            proc.WaitForExit();
            return proc.ExitCode;
        }
        catch (Exception ex)
        {
            WriteError(ex.Message);
            return 1;
        }
    }

    private static void WriteError(string message)
    {
        try
        {
            Console.Error.WriteLine("[PmAiRpaLuncher] " + message);
        }
        catch
        {
            // GUI-only session
        }
    }
}
