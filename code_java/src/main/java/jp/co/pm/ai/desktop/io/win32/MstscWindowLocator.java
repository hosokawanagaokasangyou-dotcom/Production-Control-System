package jp.co.pm.ai.desktop.io.win32;

import java.time.Duration;
import java.util.OptionalLong;
import java.util.concurrent.atomic.AtomicReference;

import com.sun.jna.platform.win32.User32;
import com.sun.jna.platform.win32.WinDef.HWND;
import com.sun.jna.platform.win32.WinUser;
import com.sun.jna.ptr.IntByReference;

import jp.co.pm.ai.desktop.io.RemoteDesktopLauncher;

/** {@code mstsc.exe} のクライアント HWND を PID から特定する。 */
public final class MstscWindowLocator {

    private static final String CLIENT_CLASS = "TscShellContainerClass";
    private static final Duration DEFAULT_TIMEOUT = Duration.ofSeconds(90);
    private static final Duration POLL = Duration.ofMillis(500);

    private MstscWindowLocator() {}

    public static OptionalLong findClientWindow(long processId) {
        return findClientWindow(processId, DEFAULT_TIMEOUT);
    }

    public static OptionalLong findClientWindow(long processId, Duration timeout) {
        if (!RemoteDesktopLauncher.isSupportedPlatform() || processId <= 0) {
            return OptionalLong.empty();
        }
        long deadline = System.nanoTime() + timeout.toNanos();
        while (System.nanoTime() < deadline) {
            OptionalLong found = scanOnce(processId);
            if (found.isPresent()) {
                return found;
            }
            try {
                Thread.sleep(POLL.toMillis());
            } catch (InterruptedException ex) {
                Thread.currentThread().interrupt();
                return OptionalLong.empty();
            }
        }
        return OptionalLong.empty();
    }

    private static OptionalLong scanOnce(long processId) {
        AtomicReference<Long> best = new AtomicReference<>();
        WinUser.WNDENUMPROC callback =
                (hWnd, lParam) -> {
                    if (!User32.INSTANCE.IsWindowVisible(hWnd)) {
                        return true;
                    }
                    char[] className = new char[256];
                    User32.INSTANCE.GetClassName(hWnd, className, className.length);
                    String cls = NativeString.fromCharArray(className);
                    if (!CLIENT_CLASS.equalsIgnoreCase(cls)) {
                        return true;
                    }
                    IntByReference pidRef = new IntByReference();
                    User32.INSTANCE.GetWindowThreadProcessId(hWnd, pidRef);
                    if (pidRef.getValue() != (int) processId) {
                        return true;
                    }
                    best.set(PointerUtil.handleToLong(hWnd));
                    return false;
                };
        User32.INSTANCE.EnumWindows(callback, null);
        Long handle = best.get();
        return handle != null && handle > 0 ? OptionalLong.of(handle) : OptionalLong.empty();
    }

    /** Win32 文字列ユーティリティ。 */
    static final class NativeString {
        private NativeString() {}

        static String fromCharArray(char[] chars) {
            int len = 0;
            while (len < chars.length && chars[len] != 0) {
                len++;
            }
            return new String(chars, 0, len);
        }
    }

    /** HWND → long。 */
    static final class PointerUtil {
        private PointerUtil() {}

        static long handleToLong(HWND hWnd) {
            if (hWnd == null || hWnd.getPointer() == null) {
                return 0L;
            }
            return com.sun.jna.Pointer.nativeValue(hWnd.getPointer());
        }
    }
}
