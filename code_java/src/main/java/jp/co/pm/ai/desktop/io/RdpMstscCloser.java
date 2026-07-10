package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Path;
import java.time.Duration;
import java.util.Optional;
import java.util.OptionalLong;
import java.util.concurrent.atomic.AtomicReference;

import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.User32;
import com.sun.jna.platform.win32.WinDef.HWND;
import com.sun.jna.platform.win32.WinUser;

import jp.co.pm.ai.desktop.io.win32.MstscWindowLocator;

/** ローカル {@code mstsc.exe} を終了する（リモートのサインアウトではない）。 */
public final class RdpMstscCloser {

    private static final Duration GRACEFUL_CLOSE_TIMEOUT = Duration.ofSeconds(12);
    private static final Duration DESTROY_TIMEOUT = Duration.ofSeconds(5);

    public enum CloseMethod {
        NOT_FOUND,
        GRACEFUL_WM_CLOSE,
        PROCESS_DESTROY,
        PROCESS_DESTROY_FORCIBLY
    }

    public record CloseResult(long processId, CloseMethod method, boolean closed) {}

    private RdpMstscCloser() {}

    /** 指定 .rdp で接続中の mstsc があれば終了する。 */
    public static CloseResult closeForProfile(Path rdpProfile) {
        if (!RemoteDesktopLauncher.isSupportedPlatform() || rdpProfile == null) {
            return new CloseResult(-1L, CloseMethod.NOT_FOUND, false);
        }
        Path abs = rdpProfile.toAbsolutePath().normalize();
        OptionalLong pid = RdpMstscProcessFinder.scanForMstscProcessId(abs);
        if (pid.isEmpty()) {
            return new CloseResult(-1L, CloseMethod.NOT_FOUND, false);
        }
        return closeProcess(pid.getAsLong());
    }

    static CloseResult closeProcess(long pid) {
        if (pid <= 0) {
            return new CloseResult(pid, CloseMethod.NOT_FOUND, false);
        }
        if (tryGracefulClose(pid)) {
            return new CloseResult(pid, CloseMethod.GRACEFUL_WM_CLOSE, true);
        }
        Optional<ProcessHandle> handle = ProcessHandle.of(pid);
        if (handle.isEmpty() || !handle.get().isAlive()) {
            return new CloseResult(pid, CloseMethod.NOT_FOUND, true);
        }
        handle.get().destroy();
        if (waitUntilExited(pid, DESTROY_TIMEOUT)) {
            return new CloseResult(pid, CloseMethod.PROCESS_DESTROY, true);
        }
        handle = ProcessHandle.of(pid);
        if (handle.isPresent() && handle.get().isAlive()) {
            handle.get().destroyForcibly();
        }
        boolean closed = waitUntilExited(pid, DESTROY_TIMEOUT);
        return new CloseResult(pid, CloseMethod.PROCESS_DESTROY_FORCIBLY, closed);
    }

    private static boolean tryGracefulClose(long pid) {
        OptionalLong hwnd = MstscWindowLocator.findSessionWindow(pid);
        if (hwnd.isEmpty()) {
            return false;
        }
        HWND window = new HWND(Pointer.createConstant(hwnd.getAsLong()));
        User32.INSTANCE.PostMessage(window, WinUser.WM_CLOSE, null, null);
        return waitUntilExited(pid, GRACEFUL_CLOSE_TIMEOUT);
    }

    private static boolean waitUntilExited(long pid, Duration timeout) {
        long deadline = System.nanoTime() + timeout.toNanos();
        while (System.nanoTime() < deadline) {
            Optional<ProcessHandle> handle = ProcessHandle.of(pid);
            if (handle.isEmpty() || !handle.get().isAlive()) {
                return true;
            }
            try {
                Thread.sleep(200L);
            } catch (InterruptedException ex) {
                Thread.currentThread().interrupt();
                return false;
            }
        }
        return ProcessHandle.of(pid).map(h -> !h.isAlive()).orElse(true);
    }

    /**
     * 起動前準備: 既存 mstsc を閉じ、監視スレッドを中断する。
     *
     * @return 閉じた PID（無ければ -1）
     */
    public static long prepareRelaunch(
            Path rdpProfile, AtomicReference<Thread> sessionWatchThread) throws IOException {
        Thread previous = sessionWatchThread != null ? sessionWatchThread.getAndSet(null) : null;
        if (previous != null) {
            previous.interrupt();
        }
        CloseResult closed = closeForProfile(rdpProfile);
        return closed.closed() ? closed.processId() : -1L;
    }
}
