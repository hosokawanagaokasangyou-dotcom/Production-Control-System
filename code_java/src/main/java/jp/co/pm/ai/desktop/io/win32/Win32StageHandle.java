package jp.co.pm.ai.desktop.io.win32;

import java.util.OptionalLong;

import javafx.stage.Window;

import jp.co.pm.ai.desktop.io.RemoteDesktopLauncher;

/** JavaFX {@link Window} から Win32 HWND を取得する。 */
public final class Win32StageHandle {

    private Win32StageHandle() {}

    public static OptionalLong resolve(Window window) {
        if (window == null || !RemoteDesktopLauncher.isSupportedPlatform()) {
            return OptionalLong.empty();
        }
        try {
            Object peer =
                    Class.forName("com.sun.javafx.stage.WindowHelper")
                            .getMethod("getPeer", Window.class)
                            .invoke(null, window);
            if (peer == null) {
                return OptionalLong.empty();
            }
            Object handle = peer.getClass().getMethod("getRawHandle").invoke(peer);
            if (handle instanceof Number number) {
                long raw = number.longValue();
                return raw != 0L ? OptionalLong.of(raw) : OptionalLong.empty();
            }
        } catch (ReflectiveOperationException | ClassCastException ignored) {
            // fall through
        }
        return OptionalLong.empty();
    }
}
