using System.Runtime.InteropServices;

namespace PmAi.RdpRemoteLauncher;

internal static class ProcessMainWindowChecker
{
    internal static bool AnyVisibleTopLevelWindow(IEnumerable<int> processIds)
    {
        var idSet = processIds as HashSet<int> ?? processIds.ToHashSet();
        if (idSet.Count == 0)
        {
            return false;
        }

        var found = false;
        EnumWindows(
            (hWnd, lParam) =>
            {
                if (!IsWindowVisible(hWnd))
                {
                    return true;
                }

                if (GetWindow(hWnd, GetWindowCmd.Owner) != nint.Zero)
                {
                    return true;
                }

                GetWindowThreadProcessId(hWnd, out var windowPid);
                if (idSet.Contains((int)windowPid))
                {
                    found = true;
                    return false;
                }

                return true;
            },
            nint.Zero);

        return found;
    }

    private static bool EnumWindows(EnumWindowsProc callback, nint lParam)
    {
        return NativeMethods.EnumWindows(callback, lParam);
    }

    private static bool IsWindowVisible(nint hWnd)
    {
        return NativeMethods.IsWindowVisible(hWnd);
    }

    private static nint GetWindow(nint hWnd, GetWindowCmd cmd)
    {
        return NativeMethods.GetWindow(hWnd, cmd);
    }

    private static uint GetWindowThreadProcessId(nint hWnd, out uint processId)
    {
        return NativeMethods.GetWindowThreadProcessId(hWnd, out processId);
    }

    private enum GetWindowCmd : uint
    {
        Owner = 4,
    }

    private delegate bool EnumWindowsProc(nint hWnd, nint lParam);

    private static class NativeMethods
    {
        [DllImport("user32.dll")]
        internal static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, nint lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsWindowVisible(nint hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern nint GetWindow(nint hWnd, GetWindowCmd uCmd);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern uint GetWindowThreadProcessId(nint hWnd, out uint lProcessId);
    }
}
