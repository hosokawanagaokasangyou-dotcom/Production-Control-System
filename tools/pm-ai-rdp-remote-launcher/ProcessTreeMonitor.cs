using System.Diagnostics;
using System.Management;

namespace PmAi.RdpRemoteLauncher;

internal sealed class ProcessTreeMonitor
{
    private readonly int _rootProcessId;
    private readonly ParsedCommand? _commandSignature;
    private readonly HashSet<int> _trackedProcessIds = new();
    private DateTime _lastStatusLogUtc = DateTime.MinValue;

    private static readonly TimeSpan PollInterval = TimeSpan.FromSeconds(2);
    private static readonly TimeSpan StatusLogInterval = TimeSpan.FromSeconds(30);

    internal ProcessTreeMonitor(int rootProcessId, ParsedCommand? commandSignature = null)
    {
        _rootProcessId = rootProcessId;
        _commandSignature = commandSignature;
        _trackedProcessIds.Add(rootProcessId);
    }

    internal void WaitUntilFinished(CancellationToken cancellationToken = default)
    {
        _lastStatusLogUtc = DateTime.UtcNow;
        LogStatus(force: true);

        while (!cancellationToken.IsCancellationRequested)
        {
            ExpandTrackedWithChildren();
            RemoveExitedProcesses();

            var treeEmpty = _trackedProcessIds.Count == 0;
            var signatureGone = !IsSignatureProcessRunning();
            if (treeEmpty && signatureGone)
            {
                LauncherLog.Info("全プロセス終了（プロセスツリー監視完了）");
                return;
            }

            LogStatusIfDue();
            cancellationToken.WaitHandle.WaitOne(PollInterval);
        }
    }

    private void ExpandTrackedWithChildren()
    {
        var snapshot = _trackedProcessIds.ToArray();
        foreach (var parentId in snapshot)
        {
            foreach (var childId in QueryChildProcessIds(parentId))
            {
                _trackedProcessIds.Add(childId);
            }
        }
    }

    private static IEnumerable<int> QueryChildProcessIds(int parentProcessId)
    {
        List<int>? childIds = null;
        try
        {
            using var searcher = new ManagementObjectSearcher(
                "SELECT ProcessId FROM Win32_Process WHERE ParentProcessId = " + parentProcessId);
            foreach (ManagementObject obj in searcher.Get())
            {
                using (obj)
                {
                    var raw = obj["ProcessId"];
                    if (raw != null && int.TryParse(raw.ToString(), out var pid) && pid > 0)
                    {
                        childIds ??= new List<int>();
                        childIds.Add(pid);
                    }
                }
            }
        }
        catch (ManagementException ex)
        {
            LauncherLog.Warn("子プロセス列挙失敗 ParentPID=" + parentProcessId + ": " + ex.Message);
        }

        return childIds ?? Enumerable.Empty<int>();
    }

    private void RemoveExitedProcesses()
    {
        var snapshot = _trackedProcessIds.ToArray();
        foreach (var pid in snapshot)
        {
            if (!IsProcessAlive(pid))
            {
                _trackedProcessIds.Remove(pid);
            }
        }
    }

    private static bool IsProcessAlive(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            return !process.HasExited;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    private bool IsSignatureProcessRunning()
    {
        if (!_commandSignature.HasValue)
        {
            return false;
        }

        return ProcessRunningChecker.IsAlreadyRunning(_commandSignature.Value, loginId: null);
    }

    private void LogStatusIfDue()
    {
        if (DateTime.UtcNow - _lastStatusLogUtc >= StatusLogInterval)
        {
            LogStatus(force: false);
        }
    }

    private void LogStatus(bool force)
    {
        if (!force && DateTime.UtcNow - _lastStatusLogUtc < StatusLogInterval)
        {
            return;
        }

        _lastStatusLogUtc = DateTime.UtcNow;
        var pids = _trackedProcessIds.Count == 0
            ? "(なし)"
            : string.Join(", ", _trackedProcessIds.OrderBy(id => id));
        var signature = IsSignatureProcessRunning() ? "あり" : "なし";
        LauncherLog.Info(
            "監視中… ルートPID="
                + _rootProcessId
                + " 追跡中="
                + _trackedProcessIds.Count
                + " ["
                + pids
                + "] 署名一致プロセス="
                + signature);
    }
}
