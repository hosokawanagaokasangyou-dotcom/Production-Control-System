package jp.co.pm.ai.desktop.io.win32;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.OptionalLong;

import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.User32;
import com.sun.jna.platform.win32.WinDef.HWND;
import com.sun.jna.ptr.IntByReference;

import jp.co.pm.ai.desktop.io.RemoteDesktopLauncher;

/** {@code mstsc.exe} セッションウィンドウ HWND の探索。 */
public final class MstscWindowLocator {

    private static final int SCORE_PID_MATCH = 100;
    private static final int SCORE_RDP_TITLE = 40;
    private static final int SCORE_CLIENT_SURFACE = 50;

    private MstscWindowLocator() {}

    public static OptionalLong findSessionWindow(long processIdHint) {
        java.util.Optional<MstscCaptureTarget> target = findCaptureTarget(processIdHint);
        if (target.isEmpty()) {
            return OptionalLong.empty();
        }
        MstscCaptureTarget t = target.get();
        long hwnd = t.frameHwnd() > 0L ? t.frameHwnd() : t.clientHwnd();
        return hwnd > 0L ? OptionalLong.of(hwnd) : OptionalLong.empty();
    }

    /** プレビュー用: 外枠 HWND を正とし、クライアントはフォールバック。 */
    public static java.util.Optional<MstscCaptureTarget> findCaptureTarget(long processIdHint) {
        if (!RemoteDesktopLauncher.isSupportedPlatform()) {
            return java.util.Optional.empty();
        }
        List<Candidate> candidates = collectTopLevelCandidates();
        Candidate best = pickBest(candidates, processIdHint);
        if (best == null) {
            return java.util.Optional.empty();
        }
        long frame = best.handle();
        long client = findClientDeep(new HWND(Pointer.createConstant(frame)));
        if (isClientSurfaceClass(best.className())) {
            return java.util.Optional.of(new MstscCaptureTarget(frame, frame));
        }
        return java.util.Optional.of(
                new MstscCaptureTarget(frame, client > 0L ? client : frame));
    }

    static boolean isClientSurfaceClass(String className) {
        return className != null && className.equalsIgnoreCase("TscShellContainerClass");
    }

    static boolean looksLikeRdpSessionTitle(String title) {
        if (title == null || title.isBlank()) {
            return false;
        }
        String t = title.toLowerCase(Locale.ROOT);
        if (t.contains("セキュリティ") || t.contains("security")) {
            return false;
        }
        return t.contains("remote desktop")
                || t.contains("リモート デスクトップ")
                || t.contains("リモートデスクトップ");
    }

    private static List<Candidate> collectTopLevelCandidates() {
        List<Candidate> out = new ArrayList<>();
        User32 user32 = User32.INSTANCE;
        user32.EnumWindows(
                (hWnd, data) -> {
                    if (isNullHwnd(hWnd) || !user32.IsWindowVisible(hWnd)) {
                        return true;
                    }
                    char[] classBuf = new char[256];
                    user32.GetClassName(hWnd, classBuf, classBuf.length);
                    String className = NativeString.fromCharArray(classBuf);
                    char[] titleBuf = new char[512];
                    user32.GetWindowText(hWnd, titleBuf, titleBuf.length);
                    String title = NativeString.fromCharArray(titleBuf);
                    IntByReference pidRef = new IntByReference();
                    user32.GetWindowThreadProcessId(hWnd, pidRef);
                    long handle = Pointer.nativeValue(hWnd.getPointer());
                    int score =
                            scoreCandidate(title, className, isClientSurfaceClass(className));
                    if (score > 0) {
                        out.add(new Candidate(handle, pidRef.getValue(), className, title, score));
                    }
                    return true;
                },
                null);
        return out;
    }

    static int scoreCandidate(String title, String className, boolean clientSurface) {
        int score = 0;
        if (clientSurface || "TscShellContainerClass".equalsIgnoreCase(className)) {
            score += SCORE_CLIENT_SURFACE;
        }
        if (looksLikeRdpSessionTitle(title)) {
            score += SCORE_RDP_TITLE;
        }
        if ("#32770".equalsIgnoreCase(className) && looksLikeRdpSessionTitle(title)) {
            score += 20;
        }
        return score;
    }

    private static Candidate pickBest(List<Candidate> candidates, long processIdHint) {
        return candidates.stream()
                .map(
                        c -> {
                            int score = c.score();
                            if (processIdHint > 0 && c.processId() == processIdHint) {
                                score += SCORE_PID_MATCH;
                            }
                            return new Candidate(
                                    c.handle(), c.processId(), c.className(), c.title(), score);
                        })
                .filter(c -> c.score() > 0)
                .max(Comparator.comparingInt(Candidate::score))
                .orElse(null);
    }

    private static long findClientDeep(HWND parent) {
        if (isNullHwnd(parent)) {
            return 0L;
        }
        IntByReference found = new IntByReference();
        User32.INSTANCE.EnumChildWindows(
                parent,
                (child, data) -> {
                    char[] classBuf = new char[256];
                    User32.INSTANCE.GetClassName(child, classBuf, classBuf.length);
                    String className = NativeString.fromCharArray(classBuf);
                    if (isClientSurfaceClass(className)) {
                        found.setValue((int) Pointer.nativeValue(child.getPointer()));
                        return false;
                    }
                    long nested = findClientDeep(child);
                    if (nested > 0L) {
                        found.setValue((int) nested);
                        return false;
                    }
                    return true;
                },
                null);
        return found.getValue() & 0xFFFFFFFFL;
    }

    private static boolean isNullHwnd(HWND hwnd) {
        return hwnd == null || hwnd.getPointer() == null || Pointer.nativeValue(hwnd.getPointer()) == 0L;
    }

    private record Candidate(
            long handle, int processId, String className, String title, int score) {}

    /** char[] → String（UTF-16 先頭 NUL まで）。 */
    private static final class NativeString {
        private NativeString() {}

        static String fromCharArray(char[] buf) {
            int len = 0;
            while (len < buf.length && buf[len] != 0) {
                len++;
            }
            return new String(buf, 0, len);
        }
    }
}
