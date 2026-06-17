package jp.co.pm.ai.desktop.io.win32;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Optional;
import java.util.OptionalLong;
import java.util.Set;

import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.User32;
import com.sun.jna.platform.win32.WinDef.HWND;
import com.sun.jna.platform.win32.WinUser;
import com.sun.jna.ptr.IntByReference;

import jp.co.pm.ai.desktop.RemoteDesktopLauncherAppIdentity;
import jp.co.pm.ai.desktop.io.RemoteDesktopLauncher;

/** {@code mstsc.exe} セッションウィンドウ HWND の探索。 */
public final class MstscWindowLocator {

    private static final int SCORE_PID_MATCH = 100;
    private static final int SCORE_RDP_TITLE = 40;
    private static final int SCORE_CLIENT_SURFACE = 50;
    private static final int SCORE_MSTSC_PROCESS = 80;

    /** プレビュー取得元から除外する自 JVM プロセス（ランチャー JavaFX 窓）。 */
    private static final long LOCAL_PROCESS_ID = ProcessHandle.current().pid();

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
            candidates = collectMstscSurfaceTreeCandidates(processIdHint);
            best = pickBest(candidates, processIdHint);
        }
        if (best == null) {
            return java.util.Optional.empty();
        }
        long surface = best.handle();
        long frame = resolveRootHwnd(surface);
        long client =
                isClientSurfaceClass(best.className())
                        ? surface
                        : findClientDeep(new HWND(Pointer.createConstant(frame)));
        if (client <= 0L) {
            client = surface;
        }
        if (isClientSurfaceClass(best.className()) && frame <= 0L) {
            frame = surface;
        }
        if (frame <= 0L) {
            frame = client;
        }
        return java.util.Optional.of(new MstscCaptureTarget(frame, client));
    }

    static boolean isClientSurfaceClass(String className) {
        if (className == null) {
            return false;
        }
        return className.equalsIgnoreCase("TscShellContainerClass")
                || className.equalsIgnoreCase("IM Client Area");
    }

    static boolean looksLikeRdpSessionTitle(String title) {
        if (title == null || title.isBlank() || isLauncherWindowTitle(title)) {
            return false;
        }
        String t = title.toLowerCase(Locale.ROOT);
        if (t.contains("セキュリティ") || t.contains("security")) {
            return false;
        }
        return t.contains("remote desktop connection")
                || t.contains("リモート デスクトップ接続")
                || t.contains("リモートデスクトップ接続");
    }

    /** プレビュー対象から除外するウィンドウ（自アプリ・ランチャー UI 等）。 */
    static boolean isExcludedCaptureWindow(String title, String className, int processId) {
        if (processId > 0 && processId == LOCAL_PROCESS_ID) {
            return true;
        }
        if (isLauncherWindowTitle(title)) {
            return true;
        }
        if (className != null) {
            String cls = className.toLowerCase(Locale.ROOT);
            if (cls.contains("glass window") || cls.equals("javafxstage")) {
                return isLauncherWindowTitle(title) || processId == LOCAL_PROCESS_ID;
            }
        }
        return false;
    }

    static boolean isLauncherWindowTitle(String title) {
        if (title == null || title.isBlank()) {
            return false;
        }
        String normalized = title.strip();
        if (normalized.equals(RemoteDesktopLauncherAppIdentity.DISPLAY_TITLE)) {
            return true;
        }
        String lower = normalized.toLowerCase(Locale.ROOT);
        return lower.contains("rpaランチャー")
                || lower.contains("rdprpalauncher")
                || lower.contains("pmairpaluncher")
                || (lower.contains("リモートデスクトップ") && lower.contains("ランチャー"));
    }

    static boolean isMstscProcessId(int processId) {
        if (processId <= 0) {
            return false;
        }
        Optional<ProcessHandle> handle = ProcessHandle.of(processId);
        if (handle.isEmpty()) {
            return false;
        }
        return handle.get()
                .info()
                .command()
                .map(cmd -> cmd.replace('/', '\\').toLowerCase(Locale.ROOT))
                .map(path -> path.endsWith("mstsc.exe"))
                .orElse(false);
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
                    int processId = pidRef.getValue();
                    if (isExcludedCaptureWindow(title, className, processId)) {
                        return true;
                    }
                    long handle = Pointer.nativeValue(hWnd.getPointer());
                    int score =
                            scoreCandidate(
                                    title, className, isClientSurfaceClass(className), processId);
                    if (score > 0) {
                        out.add(new Candidate(handle, processId, className, title, score));
                    }
                    return true;
                },
                null);
        return out;
    }

    /**
     * 全画面 mstsc はデスクトップ配下の子 HWND になることがある。
     * トップレベル探索で見つからないとき、デスクトップと各トップレベル配下を mstsc PID で再帰走査する。
     */
    private static List<Candidate> collectMstscSurfaceTreeCandidates(long processIdHint) {
        List<Candidate> out = new ArrayList<>();
        Set<Long> seen = new HashSet<>();
        User32 user32 = User32.INSTANCE;
        enumMstscSurfaceSubtree(user32.GetDesktopWindow(), processIdHint, out, seen);
        user32.EnumWindows(
                (hWnd, data) -> {
                    enumMstscSurfaceSubtree(hWnd, processIdHint, out, seen);
                    return true;
                },
                null);
        return out;
    }

    private static void enumMstscSurfaceSubtree(
            HWND root, long processIdHint, List<Candidate> out, Set<Long> seen) {
        if (isNullHwnd(root)) {
            return;
        }
        visitMstscSurfaceWindow(root, processIdHint, out, seen);
        User32.INSTANCE.EnumChildWindows(
                root,
                (child, data) -> {
                    enumMstscSurfaceSubtree(child, processIdHint, out, seen);
                    return true;
                },
                null);
    }

    private static void visitMstscSurfaceWindow(
            HWND hWnd, long processIdHint, List<Candidate> out, Set<Long> seen) {
        User32 user32 = User32.INSTANCE;
        if (isNullHwnd(hWnd) || !user32.IsWindowVisible(hWnd)) {
            return;
        }
        long handle = Pointer.nativeValue(hWnd.getPointer());
        if (!seen.add(handle)) {
            return;
        }
        char[] classBuf = new char[256];
        user32.GetClassName(hWnd, classBuf, classBuf.length);
        String className = NativeString.fromCharArray(classBuf);
        char[] titleBuf = new char[512];
        user32.GetWindowText(hWnd, titleBuf, titleBuf.length);
        String title = NativeString.fromCharArray(titleBuf);
        IntByReference pidRef = new IntByReference();
        user32.GetWindowThreadProcessId(hWnd, pidRef);
        int processId = pidRef.getValue();
        if (isExcludedCaptureWindow(title, className, processId)) {
            return;
        }
        if (processIdHint > 0 && processId != processIdHint) {
            return;
        }
        if (processIdHint <= 0 && !isMstscProcessId(processId)) {
            return;
        }
        int score =
                scoreCandidate(
                        title, className, isClientSurfaceClass(className), processId);
        if (score > 0) {
            out.add(new Candidate(handle, processId, className, title, score));
        }
    }

    static long resolveRootHwnd(long hwndNative) {
        if (hwndNative <= 0L) {
            return 0L;
        }
        HWND hwnd = new HWND(Pointer.createConstant(hwndNative));
        HWND root = User32.INSTANCE.GetAncestor(hwnd, WinUser.GA_ROOT);
        if (!isNullHwnd(root)) {
            return Pointer.nativeValue(root.getPointer());
        }
        return hwndNative;
    }

    static int scoreCandidate(String title, String className, boolean clientSurface) {
        return scoreCandidate(title, className, clientSurface, -1);
    }

    static int scoreCandidate(String title, String className, boolean clientSurface, int processId) {
        if (isExcludedCaptureWindow(title, className, processId)) {
            return 0;
        }
        boolean mstscProcess = isMstscProcessId(processId);
        int score = 0;
        if (clientSurface || isClientSurfaceClass(className)) {
            score += SCORE_CLIENT_SURFACE;
        }
        if (looksLikeRdpSessionTitle(title)) {
            score += SCORE_RDP_TITLE;
        }
        if ("#32770".equalsIgnoreCase(className) && looksLikeRdpSessionTitle(title)) {
            score += 20;
        }
        if (mstscProcess) {
            score += SCORE_MSTSC_PROCESS;
        }
        if (score > 0 && !clientSurface && !mstscProcess && !looksLikeRdpSessionTitle(title)) {
            return 0;
        }
        return score;
    }

    private static Candidate pickBest(List<Candidate> candidates, long processIdHint) {
        List<Candidate> pool = candidates;
        if (processIdHint > 0) {
            List<Candidate> pidMatched =
                    candidates.stream()
                            .filter(c -> c.processId() == processIdHint)
                            .toList();
            if (!pidMatched.isEmpty()) {
                pool = pidMatched;
            } else {
                List<Candidate> mstscOnly =
                        candidates.stream().filter(c -> isMstscProcessId(c.processId())).toList();
                if (!mstscOnly.isEmpty()) {
                    pool = mstscOnly;
                }
            }
        } else {
            List<Candidate> mstscOnly =
                    candidates.stream().filter(c -> isMstscProcessId(c.processId())).toList();
            if (!mstscOnly.isEmpty()) {
                pool = mstscOnly;
            }
        }
        return pool.stream()
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
