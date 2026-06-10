using System.Diagnostics;
using System.Management;

namespace PmAi.RdpRemoteLauncher;

internal sealed class ProcessTreeMonitor
{
    private readonly Process _rootProcess;
    private readonly int _rootProcessId;
    private readonly ParsedCommand _launchCommand;
    private readonly string? _loginId;
    private readonly bool _useSignatureCompletion;
    private readonly string? _scenarioPathFragment;
    private readonly HashSet<int> _trackedProcessIds = new();
    private bool _launchSignatureSeen;
    private DateTime? _lastVisibleWindowUtc;
    private readonly DateTime _monitorStartUtc = DateTime.UtcNow;
    private DateTime _lastStatusLogUtc = DateTime.MinValue;

    private static readonly TimeSpan PollInterval = TimeSpan.FromSeconds(2);
    private static readonly TimeSpan StatusLogInterval = TimeSpan.FromSeconds(30);
    private static readonly TimeSpan SignatureStartupGrace = TimeSpan.FromSeconds(12);
    private static readonly TimeSpan ScenarioUiIdleTimeout = TimeSpan.FromSeconds(15);

    internal ProcessTreeMonitor(Process rootProcess, ParsedCommand launchCommand, string? loginId)
    {
        _rootProcess = rootProcess ?? throw new ArgumentNullException(nameof(rootProcess));
        _rootProcessId = rootProcess.Id;
        _launchCommand = launchCommand;
        _loginId = string.IsNullOrWhiteSpace(loginId) ? null : loginId.Trim();
        _useSignatureCompletion = _loginId != null;
        _scenarioPathFragment = ProcessRunningChecker.ExtractScenarioPathFragment(launchCommand.Arguments);
        _trackedProcessIds.Add(_rootProcessId);
    }

    internal void WaitUntilFinished(CancellationToken cancellationToken = default)
    {
        _lastStatusLogUtc = DateTime.UtcNow;
        LogStatus(force: true);

        while (!cancellationToken.IsCancellationRequested)
        {
            SyncTrackedProcesses();
            RemoveExitedProcesses();
            UpdateScenarioWindowActivity();

            if (IsMonitoringComplete())
            {
                return;
            }

            LogStatusIfDue();
            cancellationToken.WaitHandle.WaitOne(PollInterval);
        }
    }

    private bool IsMonitoringComplete()
    {
        if (_useSignatureCompletion)
        {
            var running = ProcessRunningChecker.IsLaunchInstanceRunning(_launchCommand, _loginId);
            if (running)
            {
                _launchSignatureSeen = true;
                if (TryCompleteAfterScenarioUiIdle())
                {
                    return true;
                }

                return false;
            }

            RefreshRootProcessState();
            if (!_launchSignatureSeen)
            {
                // 起動直後は WMI がコマンド行を拾うまで猶予。ルートが生きていればツリー監視を継続。
                if (!_rootProcess.HasExited
                    && DateTime.UtcNow - _monitorStartUtc < SignatureStartupGrace)
                {
                    return false;
                }

                if (!_rootProcess.HasExited && _trackedProcessIds.Count > 0)
                {
                    return false;
                }
            }

            LauncherLog.Info("起動シグネチャ一致プロセスなし（監視完了）");
            return true;
        }

        if (_trackedProcessIds.Count == 0)
        {
            LauncherLog.Info("全プロセス終了（プロセスツリー監視完了）");
            return true;
        }

        return false;
    }

    private void SyncTrackedProcesses()
    {
        RefreshRootProcessState();
        if (!_rootProcess.HasExited)
        {
            _trackedProcessIds.Add(_rootProcessId);
        }

        var snapshot = _trackedProcessIds.ToArray();
        foreach (var parentId in snapshot)
        {
            foreach (var childId in QueryChildProcessIds(parentId))
            {
                _trackedProcessIds.Add(childId);
            }
        }

        if (!_useSignatureCompletion)
        {
            return;
        }

        foreach (var pid in ProcessRunningChecker.FindAllMatchingProcessIds(_launchCommand, _loginId))
        {
            _trackedProcessIds.Add(pid);
        }

        if (string.IsNullOrWhiteSpace(_scenarioPathFragment))
        {
            return;
        }

        foreach (var pid in ProcessRunningChecker.FindProcessIdsByCommandLineContains(_scenarioPathFragment))
        {
            _trackedProcessIds.Add(pid);
        }
    }

    private void RefreshRootProcessState()
    {
        try
        {
            _rootProcess.Refresh();
        }
        catch (InvalidOperationException)
        {
            // プロセス既終了
        }
    }

    private void UpdateScenarioWindowActivity()
    {
        if (string.IsNullOrWhiteSpace(_scenarioPathFragment) || !_launchSignatureSeen)
        {
            return;
        }

        var matchingProcessIds = ProcessRunningChecker.FindAllMatchingProcessIds(_launchCommand, _loginId);
        if (matchingProcessIds.Count == 0)
        {
            return;
        }

        if (ProcessMainWindowChecker.AnyVisibleTopLevelWindow(matchingProcessIds))
        {
            _lastVisibleWindowUtc = DateTime.UtcNow;
        }
    }

    private bool TryCompleteAfterScenarioUiIdle()
    {
        if (string.IsNullOrWhiteSpace(_scenarioPathFragment) || !_launchSignatureSeen || !_lastVisibleWindowUtc.HasValue)
        {
            return false;
        }

        var idleFor = DateTime.UtcNow - _lastVisibleWindowUtc.Value;
        if (idleFor < ScenarioUiIdleTimeout)
        {
            return false;
        }

        LauncherLog.Info(
            "シナリオ起動後 UI 非表示が "
                + (int)ScenarioUiIdleTimeout.TotalSeconds
                + " 秒継続（監視完了）");
        return true;
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
            process.Refresh();
            return !process.HasExited;
        }
        catch (ArgumentException)
        {
            return false;
        }
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
        var signature =
            _useSignatureCompletion
                ? (ProcessRunningChecker.IsLaunchInstanceRunning(_launchCommand, _loginId) ? "あり" : "なし")
                : "(資格情報なし)";
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
