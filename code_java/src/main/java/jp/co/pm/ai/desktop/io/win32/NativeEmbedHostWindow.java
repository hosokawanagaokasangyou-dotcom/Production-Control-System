package jp.co.pm.ai.desktop.io.win32;

import com.sun.jna.Native;
import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.User32;
import com.sun.jna.platform.win32.WinDef.HWND;
import com.sun.jna.platform.win32.WinDef.LRESULT;
import com.sun.jna.platform.win32.WinDef.WPARAM;
import com.sun.jna.platform.win32.WinDef.LPARAM;
import com.sun.jna.platform.win32.WinUser;

import jp.co.pm.ai.desktop.io.RemoteDesktopLauncher;

/** Win32 子 HWND ホスト（mstsc 埋め込み先）。 */
public final class NativeEmbedHostWindow implements AutoCloseable {

    private static final String HOST_CLASS = "PmAiRdpEmbedHost";
    private static final int WS_CHILD = 0x40000000;
    private static final int WS_VISIBLE = 0x10000000;
    private static final int WS_CLIPSIBLINGS = 0x04000000;
    private static final int WS_CLIPCHILDREN = 0x02000000;

    private static volatile boolean classRegistered;
    private static WinUser.WNDCLASSEX wndClass;

    private HWND parentHwnd;
    private HWND hostHwnd;

    public boolean create(long parentHandleNative) {
        if (!RemoteDesktopLauncher.isSupportedPlatform() || parentHandleNative == 0L) {
            return false;
        }
        ensureClassRegistered();
        parentHwnd = new HWND(Pointer.createConstant(parentHandleNative));
        hostHwnd =
                User32.INSTANCE.CreateWindowEx(
                        0,
                        HOST_CLASS,
                        "PmAiRdpEmbedHost",
                        WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN,
                        0,
                        0,
                        100,
                        100,
                        parentHwnd,
                        null,
                        null,
                        null);
        return hostHwnd != null && !Win32Pointers.isNull(hostHwnd);
    }

    public void setBounds(int x, int y, int width, int height) {
        if (hostHwnd == null || Win32Pointers.isNull(hostHwnd)) {
            return;
        }
        int w = Math.max(1, width);
        int h = Math.max(1, height);
        User32.INSTANCE.MoveWindow(hostHwnd, x, y, w, h, true);
        User32.INSTANCE.ShowWindow(hostHwnd, WinUser.SW_SHOW);
    }

    public long hostHandleNative() {
        if (hostHwnd == null || Win32Pointers.isNull(hostHwnd)) {
            return 0L;
        }
        return Pointer.nativeValue(hostHwnd.getPointer());
    }

    public HWND hostHwnd() {
        return hostHwnd;
    }

    @Override
    public void close() {
        if (hostHwnd != null && !Win32Pointers.isNull(hostHwnd)) {
            User32.INSTANCE.DestroyWindow(hostHwnd);
            hostHwnd = null;
        }
        parentHwnd = null;
    }

    private static void ensureClassRegistered() {
        if (classRegistered) {
            return;
        }
        synchronized (NativeEmbedHostWindow.class) {
            if (classRegistered) {
                return;
            }
            wndClass = new WinUser.WNDCLASSEX();
            wndClass.cbSize = wndClass.size();
            wndClass.lpfnWndProc = EmbedHostWindowProc.INSTANCE;
            wndClass.hInstance = null;
            wndClass.lpszClassName = HOST_CLASS;
            User32.INSTANCE.RegisterClassEx(wndClass);
            classRegistered = true;
        }
    }

    private abstract static class EmbedHostWindowProc implements WinUser.WindowProc {
        private static final EmbedHostWindowProc INSTANCE = new EmbedHostWindowProc() {};

        @Override
        public LRESULT callback(HWND hWnd, int uMsg, WPARAM wParam, LPARAM lParam) {
            return User32.INSTANCE.DefWindowProc(hWnd, uMsg, wParam, lParam);
        }
    }

    /** JNA HWND null 判定。 */
    static final class Win32Pointers {
        private Win32Pointers() {}

        static boolean isNull(HWND hwnd) {
            return hwnd == null || hwnd.getPointer() == null || Pointer.nativeValue(hwnd.getPointer()) == 0L;
        }
    }
}
