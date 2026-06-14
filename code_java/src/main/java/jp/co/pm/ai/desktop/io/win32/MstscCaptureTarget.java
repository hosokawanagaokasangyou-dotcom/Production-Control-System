package jp.co.pm.ai.desktop.io.win32;

/** mstsc キャプチャ試行用 HWND ペア（外枠優先）。 */
public record MstscCaptureTarget(long frameHwnd, long clientHwnd) {

    public boolean isValid() {
        return frameHwnd > 0L || clientHwnd > 0L;
    }

    /** 外枠 → クライアントの順で試す。 */
    public long[] handlesToTry() {
        if (frameHwnd > 0L && clientHwnd > 0L && frameHwnd != clientHwnd) {
            return new long[] {frameHwnd, clientHwnd};
        }
        if (frameHwnd > 0L) {
            return new long[] {frameHwnd};
        }
        if (clientHwnd > 0L) {
            return new long[] {clientHwnd};
        }
        return new long[0];
    }
}
