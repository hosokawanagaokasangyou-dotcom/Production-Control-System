using System.Text;

namespace PmAi.RdpRemoteLauncher;

internal sealed class LauncherIni
{
    internal const string SelectedSlotKey = "起動プログラム番号";
    internal const string DefaultIniFileName = "RAP設定.ini";
    internal const string IniPathEnvVar = "PM_AI_RDP_LAUNCHER_INI";

    internal int SelectedSlot { get; set; } = 1;

    internal Dictionary<int, string> Slots { get; } = new();

    internal static LauncherIni Load(string path)
    {
        var ini = new LauncherIni();
        foreach (var rawLine in File.ReadAllLines(path, Encoding.UTF8))
        {
            var line = rawLine.Trim();
            if (line.Length == 0 || line.StartsWith('#') || line.StartsWith(';'))
            {
                continue;
            }

            var eq = line.IndexOf('=');
            if (eq <= 0)
            {
                continue;
            }

            var key = line[..eq].Trim();
            var value = line[(eq + 1)..].Trim();
            if (key == SelectedSlotKey)
            {
                if (int.TryParse(value, out var slot) && slot >= 1)
                {
                    ini.SelectedSlot = slot;
                }
                continue;
            }

            if (int.TryParse(key, out var slotNumber) && slotNumber >= 1 && !string.IsNullOrWhiteSpace(value))
            {
                ini.Slots[slotNumber] = value;
            }
        }

        return ini;
    }

    internal string? ResolveSelectedCommand()
    {
        return Slots.TryGetValue(SelectedSlot, out var command) ? command : null;
    }
}
