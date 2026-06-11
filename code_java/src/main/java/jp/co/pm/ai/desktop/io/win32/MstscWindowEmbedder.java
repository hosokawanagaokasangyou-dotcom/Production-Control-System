package jp.co.pm.ai.desktop.io.win32;

import java.util.OptionalLong;

import com.sun.jna.platform.win32.User32;
import com.sun.jna.platform.win32.WinDef.HWND;
import com.sun.jna.platform.win32.WinUser;

import javafx.geometry.Bounds;
import javafx.scene.Node;
import javafx.stage.Window;

import jp.co.pm.ai.desktop.io.RemoteDesktopLauncher;

/**
 * mstsc クライアント HWND を JavaFX ノード領域へ reparent する。
 *
 * <p>Windows 専用。非 Windows では attach が false を返す。
 */
public final class MstscWindowEmbedder implements AutoCloseable {

    private static final int GWL_STYLE = -16;
    private static final int WS_CHILD = 0x40000000;
    private static final int WS_VISIBLE = 0x10000000;
    private static final int WS_CAPTION = 0x00C00000;
    private static final int WS_THICKFRAME = 0x00040000;
    private static final int WS_POPUP = uncheckedInt(0x80000000L);
    private static final int SWP_NOZORDER = 0x0004;
    private static final int SWP_FRAMECHANGED = 0x0020;
    private static final HWND HWND_TOP = new HWND(com.sun.jna.Pointer.createConstant(0));

    private final NativeEmbedHostWindow host = new NativeEmbedHostWindow();
    private long mstscHwnd;
    private Window boundWindow;
    private Node boundNode;

    public boolean isSupported() {
        return RemoteDesktopLauncher.isSupportedPlatform();
    }

    public boolean isAttached() {
        return mstscHwnd != 0L && host.hostHandleNative() != 0L;
    }

    /**
     * 子ホストを作成し mstsc を reparent する。
     *
     * @return 成功した場合 {@code true}
     */
    public boolean attach(Window window, Node anchor, long mstscHwndNative) {
        if (!isSupported() || window == null || anchor == null || mstscHwndNative == 0L) {
            return false;
        }
        OptionalLong stageHandle = Win32StageHandle.resolve(window);
        if (stageHandle.isEmpty()) {
            return false;
        }
        if (!host.create(stageHandle.getAsLong())) {
            return false;
        }
        boundWindow = window;
        boundNode = anchor;
        syncBounds();

        HWND client = new HWND(com.sun.jna.Pointer.createConstant(mstscHwndNative));
        HWND hostHwnd = host.hostHwnd();
        if (NativeEmbedHostWindow.Win32Pointers.isNull(hostHwnd)) {
            detach();
            return false;
        }

        User32 user32 = User32.INSTANCE;
        user32.SetParent(client, hostHwnd);

        int style = user32.GetWindowLong(client, GWL_STYLE);
        style = (style & ~WS_POPUP & ~WS_CAPTION & ~WS_THICKFRAME) | WS_CHILD | WS_VISIBLE;
        user32.SetWindowLong(client, GWL_STYLE, style);
        user32.SetWindowPos(
                client,
                HWND_TOP,
                0,
                0,
                Math.max(1, (int) anchor.getLayoutBounds().getWidth()),
                Math.max(1, (int) anchor.getLayoutBounds().getHeight()),
                SWP_NOZORDER | SWP_FRAMECHANGED);

        mstscHwnd = mstscHwndNative;
        return true;
    }

    /** アンカー Node の現在 bounds にホストと mstsc を合わせる。 */
    public void syncBounds() {
        if (boundWindow == null || boundNode == null || host.hostHandleNative() == 0L) {
            return;
        }
        Bounds anchorBounds = boundNode.localToScene(boundNode.getLayoutBounds());
        Bounds stageBounds = boundWindow.getScene().getRoot().localToScene(boundWindow.getScene().getRoot().getLayoutBounds());
        int x = (int) Math.round(anchorBounds.getMinX() - stageBounds.getMinX());
        int y = (int) Math.round(anchorBounds.getMinY() - stageBounds.getMinY());
        int width = (int) Math.round(anchorBounds.getWidth());
        int height = (int) Math.round(anchorBounds.getHeight());
        host.setBounds(x, y, width, height);

        if (mstscHwnd != 0L) {
            HWND client = new HWND(com.sun.jna.Pointer.createConstant(mstscHwnd));
            User32.INSTANCE.MoveWindow(client, 0, 0, Math.max(1, width), Math.max(1, height), true);
        }
    }

    /** reparent を解除し子ホストを破棄する。 */
    public void detach() {
        if (mstscHwnd != 0L && isSupported()) {
            HWND client = new HWND(com.sun.jna.Pointer.createConstant(mstscHwnd));
            User32.INSTANCE.SetParent(client, null);
        }
        mstscHwnd = 0L;
        boundNode = null;
        boundWindow = null;
        host.close();
    }

    @Override
    public void close() {
        detach();
    }

    private static int uncheckedInt(long value) {
        return (int) value;
    }
}
