package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.Optional;
import java.util.OptionalLong;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;

/**
 * ローカル {@code mstsc.exe} の終了を監視する。
 *
 * <p>接続先 {@code PmAiRdpRemoteLauncher} が子プロセス終了後に RDP を切断すると mstsc も終了するため、
 * 配台システム側で RPA 完了相当のタイミングを検知できる。
 */
public final class RdpMstscSessionMonitor {

    /** 監視結果。 */
    public record SessionEndEvent(Path rdpProfile, int exitCode, EndReason reason) {}

    public enum EndReason {
        /** mstsc プロセスが終了した。 */
        MSTSC_EXIT,
        /** 起動直後に mstsc を特定できなかった。 */
        PROCESS_NOT_FOUND,
        /** 監視中に割り込み等。 */
        MONITOR_INTERRUPTED
    }

    private static final Duration FIND_TIMEOUT = Duration.ofSeconds(90);
    private static final Duration PID_MARKER_TIMEOUT = Duration.ofSeconds(20);
    private static final Duration POLL_INTERVAL = Duration.ofMillis(500);

    private RdpMstscSessionMonitor() {}

    /**
     * バックグラウンドで mstsc 終了を待ち、終了時に {@code onEnded} を呼ぶ（FX スレッドではない）。
     *
     * @param cancelPrevious 前回の監視スレッド（あれば interrupt）
     */
    public static void watchAfterLaunch(
            Path rdpProfile,
            OptionalLong knownMstscPid,
            Optional<Path> mstscPidMarkerFile,
            AtomicReference<Thread> cancelPrevious,
            Consumer<SessionEndEvent> onEnded) {
        if (!RemoteDesktopLauncher.isSupportedPlatform()) {
            return;
        }
        Thread previous = cancelPrevious.getAndSet(null);
        if (previous != null) {
            previous.interrupt();
        }
        Thread worker =
                new Thread(
                        () -> runWatch(rdpProfile, knownMstscPid, mstscPidMarkerFile, onEnded),
                        "rdp-mstsc-watch");
        worker.setDaemon(true);
        cancelPrevious.set(worker);
        worker.start();
    }

    private static void runWatch(
            Path rdpProfile,
            OptionalLong knownMstscPid,
            Optional<Path> mstscPidMarkerFile,
            Consumer<SessionEndEvent> onEnded) {
        Path markerFile = null;
        try {
            Path abs = RemoteDesktopLauncher.validateRdpProfile(rdpProfile);
            markerFile = mstscPidMarkerFile.orElse(null);
            long pid = resolveMstscPid(abs, knownMstscPid, markerFile);
            if (pid < 0) {
                onEnded.accept(new SessionEndEvent(abs, -1, EndReason.PROCESS_NOT_FOUND));
                return;
            }
            Optional<ProcessHandle> handle = ProcessHandle.of(pid);
            if (handle.isEmpty() || !handle.get().isAlive()) {
                onEnded.accept(new SessionEndEvent(abs, 0, EndReason.MSTSC_EXIT));
                return;
            }
            int exitCode = 0;
            handle.get().onExit().get();
            if (Thread.currentThread().isInterrupted()) {
                onEnded.accept(new SessionEndEvent(abs, exitCode, EndReason.MONITOR_INTERRUPTED));
                return;
            }
            onEnded.accept(new SessionEndEvent(abs, exitCode, EndReason.MSTSC_EXIT));
        } catch (InterruptedException ex) {
            Thread.currentThread().interrupt();
            try {
                Path abs = RemoteDesktopLauncher.validateRdpProfile(rdpProfile);
                onEnded.accept(new SessionEndEvent(abs, -1, EndReason.MONITOR_INTERRUPTED));
            } catch (IOException ignored) {
                // ignore
            }
        } catch (Exception ex) {
            try {
                Path abs = RemoteDesktopLauncher.validateRdpProfile(rdpProfile);
                onEnded.accept(new SessionEndEvent(abs, -1, EndReason.PROCESS_NOT_FOUND));
            } catch (IOException ioEx) {
                onEnded.accept(
                        new SessionEndEvent(
                                rdpProfile.toAbsolutePath().normalize(),
                                -1,
                                EndReason.PROCESS_NOT_FOUND));
            }
        } finally {
            deleteQuietly(markerFile);
        }
    }

    private static long resolveMstscPid(
            Path rdpProfile, OptionalLong knownMstscPid, Path markerFile)
            throws InterruptedException {
        if (knownMstscPid.isPresent()) {
            long pid = knownMstscPid.getAsLong();
            Optional<ProcessHandle> handle = ProcessHandle.of(pid);
            if (handle.isPresent() && handle.get().isAlive()) {
                return pid;
            }
        }
        if (markerFile != null) {
            OptionalLong fromMarker = RdpMstscProcessFinder.pollPidMarkerFile(markerFile, PID_MARKER_TIMEOUT);
            if (fromMarker.isPresent()) {
                return fromMarker.getAsLong();
            }
        }
        return findMstscProcessId(rdpProfile, FIND_TIMEOUT);
    }

    static long findMstscProcessId(Path rdpProfile, Duration timeout) throws InterruptedException {
        Path abs = rdpProfile.toAbsolutePath().normalize();
        long deadline = System.nanoTime() + timeout.toNanos();
        while (System.nanoTime() < deadline) {
            if (Thread.currentThread().isInterrupted()) {
                throw new InterruptedException();
            }
            OptionalLong pid = RdpMstscProcessFinder.scanForMstscProcessId(abs);
            if (pid.isPresent()) {
                return pid.getAsLong();
            }
            Thread.sleep(POLL_INTERVAL.toMillis());
        }
        return -1L;
    }

    static boolean isMstscProcess(ProcessHandle handle) {
        return handle.info()
                .command()
                .map(cmd -> cmd.replace('/', '\\').toLowerCase(java.util.Locale.ROOT))
                .map(path -> path.endsWith("mstsc.exe"))
                .orElse(false);
    }

    static boolean processRefersToProfile(ProcessHandle handle, Path rdpProfile) {
        Path target = rdpProfile.toAbsolutePath().normalize();
        Optional<String[]> args = handle.info().arguments();
        if (args.isPresent()) {
            for (String arg : args.get()) {
                if (arg == null || arg.isBlank()) {
                    continue;
                }
                if (profileMatchKey(Path.of(arg)).equalsIgnoreCase(profileMatchKey(target))) {
                    return true;
                }
            }
        }
        return false;
    }

    static String profileMatchKey(Path path) {
        String raw = path.toString().replace('/', '\\');
        if (looksLikeWindowsAbsolutePath(raw)) {
            return raw;
        }
        return path.toAbsolutePath().normalize().toString().replace('/', '\\');
    }

    private static boolean looksLikeWindowsAbsolutePath(String path) {
        return path.length() >= 3
                && Character.isLetter(path.charAt(0))
                && path.charAt(1) == ':'
                && (path.charAt(2) == '\\' || path.charAt(2) == '/');
    }

    private static void deleteQuietly(Path markerFile) {
        if (markerFile == null) {
            return;
        }
        try {
            Files.deleteIfExists(markerFile);
        } catch (IOException ignored) {
            // ignore
        }
    }
}
