package jp.co.pm.ai.desktop.io;

import java.nio.file.Path;
import java.time.Duration;
import java.util.OptionalLong;

import jp.co.pm.ai.desktop.io.win32.MstscCaptureTarget;
import jp.co.pm.ai.desktop.io.win32.MstscWindowCapture;
import jp.co.pm.ai.desktop.io.win32.MstscWindowLocator;

/** mstsc 右ペイン読み取り専用プレビュー向け PID / HWND 解決。 */
public final class RdpMstscPreviewSupport {

    private static final Duration PID_MARKER_TIMEOUT = Duration.ofSeconds(20);
    private static final Duration SCAN_TIMEOUT = Duration.ofSeconds(90);

    private RdpMstscPreviewSupport() {}

    public static long tryResolveMstscPid(
            Path rdpProfile, OptionalLong knownPid, Path pidMarkerFile) {
        try {
            return RdpMstscProcessFinder.resolveMstscPid(
                    rdpProfile,
                    knownPid,
                    pidMarkerFile,
                    PID_MARKER_TIMEOUT,
                    SCAN_TIMEOUT);
        } catch (InterruptedException ex) {
            Thread.currentThread().interrupt();
            return -1L;
        }
    }

    public static OptionalLong findSessionWindowOnce(long processIdHint) {
        return MstscWindowLocator.findSessionWindow(processIdHint);
    }

    public static java.util.Optional<MstscCaptureTarget> findCaptureTargetOnce(long processIdHint) {
        return MstscWindowLocator.findCaptureTarget(processIdHint);
    }

    public static boolean isCaptureSupported() {
        return MstscWindowCapture.isSupported();
    }
}
