package jp.co.pm.ai.desktop.io;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.OptionalLong;
import java.util.concurrent.TimeUnit;

/** Windows 上で {@code mstsc.exe} プロセスを .rdp パスから特定する。 */
final class RdpMstscProcessFinder {

    private static final Duration POLL = Duration.ofMillis(500);

    private RdpMstscProcessFinder() {}

    static OptionalLong readPidMarkerFile(Path markerFile) {
        if (markerFile == null || !Files.isRegularFile(markerFile)) {
            return OptionalLong.empty();
        }
        try {
            String raw = Files.readString(markerFile, StandardCharsets.UTF_8).trim();
            if (raw.isEmpty()) {
                return OptionalLong.empty();
            }
            long pid = Long.parseLong(raw);
            return pid > 0 ? OptionalLong.of(pid) : OptionalLong.empty();
        } catch (IOException | NumberFormatException ex) {
            return OptionalLong.empty();
        }
    }

    static OptionalLong pollPidMarkerFile(Path markerFile, Duration timeout)
            throws InterruptedException {
        if (markerFile == null) {
            return OptionalLong.empty();
        }
        long deadline = System.nanoTime() + timeout.toNanos();
        while (System.nanoTime() < deadline) {
            if (Thread.currentThread().isInterrupted()) {
                throw new InterruptedException();
            }
            OptionalLong pid = readPidMarkerFile(markerFile);
            if (pid.isPresent()) {
                return pid;
            }
            Thread.sleep(200);
        }
        return OptionalLong.empty();
    }

    /**
     * 起動直後の mstsc PID を解決する（マーカーファイル待ち・プロセススキャンを含む）。
     *
     * <p>セキュリティダイアログ自動操作経由起動では PID マーカー書き込みまで遅延するため、
     * {@link RemoteDesktopLauncher#launch} 直後は空になり得る。
     */
    static long resolveMstscPid(
            Path rdpProfile,
            OptionalLong knownPid,
            Path markerFile,
            Duration markerTimeout,
            Duration scanTimeout)
            throws InterruptedException {
        if (knownPid.isPresent()) {
            long pid = knownPid.getAsLong();
            if (isAlive(pid)) {
                return pid;
            }
        }
        if (markerFile != null) {
            OptionalLong fromMarker = pollPidMarkerFile(markerFile, markerTimeout);
            if (fromMarker.isPresent()) {
                return fromMarker.getAsLong();
            }
        }
        Path abs = rdpProfile != null ? rdpProfile.toAbsolutePath().normalize() : null;
        if (abs == null) {
            return -1L;
        }
        long deadline = System.nanoTime() + scanTimeout.toNanos();
        while (System.nanoTime() < deadline) {
            if (Thread.currentThread().isInterrupted()) {
                throw new InterruptedException();
            }
            OptionalLong pid = scanForMstscProcessId(abs);
            if (pid.isPresent()) {
                return pid.getAsLong();
            }
            Thread.sleep(POLL.toMillis());
        }
        return -1L;
    }

    /** 1 回だけ PID 解決を試みる（ポーリングループ内用）。 */
    static long tryResolveMstscPid(
            Path rdpProfile, OptionalLong knownPid, Path markerFile) {
        if (knownPid.isPresent()) {
            long pid = knownPid.getAsLong();
            if (isAlive(pid)) {
                return pid;
            }
        }
        OptionalLong fromMarker = readPidMarkerFile(markerFile);
        if (fromMarker.isPresent()) {
            return fromMarker.getAsLong();
        }
        if (rdpProfile != null) {
            OptionalLong scanned = scanForMstscProcessId(rdpProfile.toAbsolutePath().normalize());
            if (scanned.isPresent()) {
                return scanned.getAsLong();
            }
        }
        return -1L;
    }

    static OptionalLong scanForMstscProcessId(Path rdpProfile) {
        OptionalLong fromHandle = scanViaProcessHandle(rdpProfile);
        if (fromHandle.isPresent()) {
            return fromHandle;
        }
        if (!RemoteDesktopLauncher.isSupportedPlatform()) {
            return OptionalLong.empty();
        }
        OptionalLong fromWmi = scanViaWindowsWmi(rdpProfile);
        if (fromWmi.isPresent()) {
            return fromWmi;
        }
        return scanSingleMstscFallback();
    }

    static boolean commandLineRefersToProfile(String commandLine, Path rdpProfile) {
        if (commandLine == null || commandLine.isBlank()) {
            return false;
        }
        String cmd = commandLine.replace('/', '\\');
        String target = RdpMstscSessionMonitor.profileMatchKey(rdpProfile);
        if (cmd.equalsIgnoreCase(target)) {
            return true;
        }
        if (cmd.toLowerCase(Locale.ROOT).contains(target.toLowerCase(Locale.ROOT))) {
            return true;
        }
        int lastSep = Math.max(target.lastIndexOf('\\'), target.lastIndexOf('/'));
        String fileName = lastSep >= 0 ? target.substring(lastSep + 1) : target;
        String fileLower = fileName.toLowerCase(Locale.ROOT);
        String cmdLower = cmd.toLowerCase(Locale.ROOT);
        if (cmdLower.contains("\\" + fileLower) || cmdLower.contains("\"" + fileLower + "\"")) {
            return true;
        }
        return cmdLower.endsWith(fileLower)
                || cmdLower.contains(" " + fileLower)
                || cmdLower.contains("\t" + fileLower);
    }

    private static boolean isAlive(long pid) {
        if (pid <= 0) {
            return false;
        }
        return ProcessHandle.of(pid).map(ProcessHandle::isAlive).orElse(false);
    }

    private static OptionalLong scanViaProcessHandle(Path rdpProfile) {
        for (ProcessHandle handle : ProcessHandle.allProcesses().toList()) {
            if (!RdpMstscSessionMonitor.isMstscProcess(handle)) {
                continue;
            }
            if (RdpMstscSessionMonitor.processRefersToProfile(handle, rdpProfile)) {
                return OptionalLong.of(handle.pid());
            }
        }
        return OptionalLong.empty();
    }

    private static OptionalLong scanViaWindowsWmi(Path rdpProfile) {
        try {
            Process process =
                    new ProcessBuilder(
                                    "powershell.exe",
                                    "-NoProfile",
                                    "-NonInteractive",
                                    "-Command",
                                    "Get-CimInstance Win32_Process -Filter \"Name='mstsc.exe'\" "
                                            + "| ForEach-Object { $_.ProcessId.ToString() + \"`t\" + $_.CommandLine }")
                            .redirectErrorStream(true)
                            .start();
            if (!process.waitFor(8, TimeUnit.SECONDS)) {
                process.destroyForcibly();
                return OptionalLong.empty();
            }
            List<String> lines = new ArrayList<>();
            try (BufferedReader reader =
                    new BufferedReader(
                            new InputStreamReader(process.getInputStream(), StandardCharsets.UTF_8))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    if (!line.isBlank()) {
                        lines.add(line);
                    }
                }
            }
            OptionalLong matched = OptionalLong.empty();
            for (String line : lines) {
                int tab = line.indexOf('\t');
                if (tab <= 0) {
                    continue;
                }
                String pidRaw = line.substring(0, tab).trim();
                String commandLine = line.substring(tab + 1).trim();
                if (!commandLineRefersToProfile(commandLine, rdpProfile)) {
                    continue;
                }
                try {
                    long pid = Long.parseLong(pidRaw);
                    matched = OptionalLong.of(pid);
                    break;
                } catch (NumberFormatException ignored) {
                    // try next
                }
            }
            return matched;
        } catch (IOException | InterruptedException ex) {
            if (ex instanceof InterruptedException) {
                Thread.currentThread().interrupt();
            }
            return OptionalLong.empty();
        }
    }

    /** 単一 mstsc のみ稼働中ならそれを採用（コマンド行取得不可時の最終手段）。 */
    private static OptionalLong scanSingleMstscFallback() {
        OptionalLong first = OptionalLong.empty();
        int count = 0;
        for (ProcessHandle handle : ProcessHandle.allProcesses().toList()) {
            if (!RdpMstscSessionMonitor.isMstscProcess(handle)) {
                continue;
            }
            count++;
            first = OptionalLong.of(handle.pid());
            if (count > 1) {
                return OptionalLong.empty();
            }
        }
        return count == 1 ? first : OptionalLong.empty();
    }
}
