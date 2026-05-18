package jp.co.pm.ai.desktop.debug;

import java.lang.management.ManagementFactory;
import java.lang.management.MemoryUsage;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Comparator;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * NDJSON append for Cursor debug sessions. Resolves log path across Windows / WSL / fat JAR layouts.
 *
 * <p>Project convention: {@code .cursor/rules/agent-debug-ndjson-logging.mdc}（モノレポでは親
 * {@code 工程管理AIプロジェクト_JAVA/.cursor/rules/agent-debug-ndjson-logging.mdc} が正本）。
 *
 * <p>Resolution order (first hit wins):
 *
 * <ol>
 *   <li>{@code CURSOR_DEBUG_LOG} or {@code PM_AI_DEBUG_LOG} (absolute path to log file)</li>
 *   <li>UI env {@link AppPaths#KEY_PM_AI_CURSOR_DEBUG_LOG}</li>
 *   <li>{@code parent(resolveRepoRoot) / .cursor / debug-&lt;sessionId&gt;.log} (workspace {@code .cursor} when repo is
 *       {@code Production-Control-System})</li>
 *   <li>{@code resolveRepoRoot / .cursor / debug-&lt;sessionId&gt;.log}</li>
 * </ol>
 *
 * <p>If the primary path cannot be written, retries under {@code user.home/.cursor/}.
 *
 * <p>WSL コンソールから {@code powershell.exe} で Windows JVM を起動した場合、NDJSON はドライブレター経由（{@code C:\...}）で
 * 書かれることがある。Cursor がワークスペースを {@code /mnt/c/...}（WSL）として開いていると、エージェントがログを見つけにくい。
 * そのため一次書き込み成功後にミラー追記する。詳細は {@code .cursor/rules/agent-debug-wsl-windows-mirror.mdc}。
 */
public final class AgentDebugLog {

    public static final String DEFAULT_SESSION_ID = "e04a1d";

    private static final ObjectMapper JSON = new ObjectMapper();

    /** Lazily resolved when {@code WSL_DISTRO_NAME} / {@code PM_AI_WSL_DISTRO} are unset (powershell.exe 子プロセスではよく未設定）。 */
    private static volatile boolean wslDistroDiscoveryAttempted;

    private static volatile String cachedDiscoveredWslDistro;

    private AgentDebugLog() {}

    /** Heap used/max in MiB for debug NDJSON payloads (session tooling). */
    public static Map<String, Object> debugHeapMap() {
        MemoryUsage h = ManagementFactory.getMemoryMXBean().getHeapMemoryUsage();
        Map<String, Object> m = new LinkedHashMap<>();
        m.put("heapUsedMiB", h.getUsed() / (1024L * 1024L));
        long max = h.getMax();
        m.put("heapMaxMiB", max > 0 ? max / (1024L * 1024L) : -1L);
        return m;
    }

    public static Path resolveNdjsonPath(Map<String, String> ui, String sessionId) {
        String id =
                sessionId == null || sessionId.isBlank()
                        ? DEFAULT_SESSION_ID
                        : sessionId.trim();
        String fileName = "debug-" + id + ".log";

        Path envPath = pathFromEnvOverride();
        if (envPath != null) {
            return envPath;
        }

        Map<String, String> u = ui != null ? ui : Map.of();
        String uiDbg = trim(u.get("PM_AI_DEBUG_LOG"));
        if (!uiDbg.isEmpty()) {
            return Path.of(uiDbg).toAbsolutePath().normalize();
        }
        String uiPath = trim(u.get(AppPaths.KEY_PM_AI_CURSOR_DEBUG_LOG));
        if (!uiPath.isEmpty()) {
            return Path.of(uiPath).toAbsolutePath().normalize();
        }

        Path repo = AppPaths.resolveRepoRoot(u);
        /*
         * Nested clone: repo leaf is Production-Control-System → workspace .cursor is parent(repo)/.cursor
         * (see agent-debug-ndjson-logging.mdc). Flat repo: repo/.cursor (parent would be drive root — wrong).
         */
        Path parent = repo.getParent();
        if (parent != null && isProductionControlSystemRepoLeaf(repo)) {
            return parent.resolve(".cursor").resolve(fileName).toAbsolutePath().normalize();
        }
        return repo.resolve(".cursor").resolve(fileName).toAbsolutePath().normalize();
    }

    private static boolean isProductionControlSystemRepoLeaf(Path repo) {
        if (repo == null) {
            return false;
        }
        Path leaf = repo.getFileName();
        return leaf != null
                && "Production-Control-System".equalsIgnoreCase(leaf.toString());
    }

    /**
     * Appends one UTF-8 line; creates parent directories. Falls back to {@code user.home/.cursor/} only if primary
     * fails.
     *
     * @return file path actually written, or {@code null} if both attempts failed
     */
    public static Path appendNdjsonLine(Map<String, String> ui, String sessionId, String jsonLine) {
        String line = jsonLine.endsWith("\n") ? jsonLine : jsonLine + "\n";
        Path primary = resolveNdjsonPath(ui, sessionId);
        if (writeUtf8Append(primary, line)) {
            appendMirrors(primary, line, ui);
            return primary;
        }
        String id =
                sessionId == null || sessionId.isBlank()
                        ? DEFAULT_SESSION_ID
                        : sessionId.trim();
        Path fallback =
                Path.of(System.getProperty("user.home", "."))
                        .resolve(".cursor")
                        .resolve("debug-" + id + ".log")
                        .toAbsolutePath()
                        .normalize();
        if (writeUtf8Append(fallback, line)) {
            appendMirrors(fallback, line, ui);
            return fallback;
        }
        return null;
    }

    /** Package-private: mirror path string for tests ({@code C:\...} → {@code \\wsl$\distro\mnt\c\...}). */
    static String buildWslUncPathString(String windowsNormalizedAbsolute, String distro) {
        return buildWslUncPathString(windowsNormalizedAbsolute, distro, "\\\\wsl$\\");
    }

    /**
     * @param uncRootPrefix UNC のルート（{@code \\\\wsl$\\} または {@code \\\\wsl.localhost\\}）。末尾はバックスラッシュを付ける。
     */
    static String buildWslUncPathString(
            String windowsNormalizedAbsolute, String distro, String uncRootPrefix) {
        if (distro == null || distro.isBlank()) {
            return null;
        }
        if (windowsNormalizedAbsolute == null || windowsNormalizedAbsolute.length() < 3) {
            return null;
        }
        char dl = windowsNormalizedAbsolute.charAt(0);
        if (!Character.isLetter(dl) || windowsNormalizedAbsolute.charAt(1) != ':') {
            return null;
        }
        String tail = windowsNormalizedAbsolute.substring(2).replace('/', '\\');
        if (!tail.startsWith("\\")) {
            tail = "\\" + tail;
        }
        String root = uncRootPrefix != null ? uncRootPrefix : "\\\\wsl$\\";
        return root
                + distro.trim()
                + "\\mnt\\"
                + Character.toLowerCase(dl)
                + tail;
    }

    private static boolean isWindowsOs() {
        return System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("windows");
    }

    private static boolean wslUncMirrorEnabled() {
        String v = trim(System.getenv("PM_AI_DEBUG_LOG_WSL_UNC"));
        if (v.isEmpty()) {
            return true;
        }
        return !(v.equals("0") || v.equalsIgnoreCase("false") || v.equalsIgnoreCase("off"));
    }

    private static String resolveWslDistroName() {
        String d = trim(System.getenv("PM_AI_WSL_DISTRO"));
        if (!d.isEmpty()) {
            return d;
        }
        d = trim(System.getenv("WSL_DISTRO_NAME"));
        if (!d.isEmpty()) {
            return d;
        }
        return discoverWslDistroNameCached();
    }

    private static String discoverWslDistroNameCached() {
        if (wslDistroDiscoveryAttempted) {
            return cachedDiscoveredWslDistro;
        }
        synchronized (AgentDebugLog.class) {
            if (wslDistroDiscoveryAttempted) {
                return cachedDiscoveredWslDistro;
            }
            wslDistroDiscoveryAttempted = true;
            cachedDiscoveredWslDistro = discoverWslDistroByListingShare();
            return cachedDiscoveredWslDistro;
        }
    }

    /**
     * {@code \\\\wsl$\\} を列挙し、docker-desktop 等を除いて優先する（名前に {@code ubuntu} を含むものを優先）。
     */
    private static String discoverWslDistroByListingShare() {
        if (!isWindowsOs()) {
            return null;
        }
        try (Stream<Path> stream = Files.list(Path.of("\\\\wsl$\\"))) {
            List<String> names =
                    stream.filter(
                                    p -> {
                                        try {
                                            return Files.isDirectory(p);
                                        } catch (Throwable ignored) {
                                            return false;
                                        }
                                    })
                            .map(p -> p.getFileName().toString())
                            .filter(AgentDebugLog::isPlausibleUserWslDistroFolder)
                            .sorted()
                            .collect(Collectors.toList());
            Optional<String> ubuntu =
                    names.stream()
                            .filter(n -> n.toLowerCase(Locale.ROOT).contains("ubuntu"))
                            .min(Comparator.naturalOrder());
            return ubuntu.orElseGet(() -> names.isEmpty() ? null : names.get(0));
        } catch (Throwable ignored) {
            return null;
        }
    }

    private static boolean isPlausibleUserWslDistroFolder(String name) {
        if (name == null || name.isBlank()) {
            return false;
        }
        String lower = name.toLowerCase(Locale.ROOT);
        if (lower.equals("docker-desktop")
                || lower.startsWith("rancher-desktop")
                || lower.contains("podman")) {
            return false;
        }
        return true;
    }

    private static Path windowsDrivePathToWslUnc(Path absoluteDrivePath, String distro, String uncRoot) {
        String unc = buildWslUncPathString(absoluteDrivePath.normalize().toString(), distro, uncRoot);
        return unc != null ? Path.of(unc) : null;
    }

    private static void appendMirrors(Path writtenPath, String line, Map<String, String> ui) {
        for (Path mirror : resolveMirrorTargets(writtenPath, ui)) {
            if (mirror == null) {
                continue;
            }
            Path n = mirror.toAbsolutePath().normalize();
            Path w = writtenPath.toAbsolutePath().normalize();
            if (n.equals(w)) {
                continue;
            }
            try {
                if (Files.exists(w) && Files.exists(n) && Files.isSameFile(w, n)) {
                    continue;
                }
            } catch (Throwable ignored) {
                // attempt mirror write
            }
            writeUtf8Append(n, line);
        }
    }

    private static List<Path> resolveMirrorTargets(Path primaryWritten, Map<String, String> ui) {
        Set<Path> out = new LinkedHashSet<>();
        Path explicit = mirrorPathFromEnv(ui);
        if (explicit != null) {
            out.add(explicit);
        }
        if (isWindowsOs() && wslUncMirrorEnabled()) {
            String distro = resolveWslDistroName();
            if (distro != null) {
                Path uncWsl = windowsDrivePathToWslUnc(primaryWritten, distro, "\\\\wsl$\\");
                if (uncWsl != null) {
                    out.add(uncWsl);
                }
                Path uncLh = windowsDrivePathToWslUnc(primaryWritten, distro, "\\\\wsl.localhost\\");
                if (uncLh != null) {
                    out.add(uncLh);
                }
            }
        }
        return new ArrayList<>(out);
    }

    private static Path mirrorPathFromEnv(Map<String, String> ui) {
        String v = trim(System.getenv("PM_AI_DEBUG_LOG_MIRROR"));
        if (v.isEmpty() && ui != null) {
            v = trim(ui.get(AppPaths.KEY_PM_AI_DEBUG_LOG_MIRROR));
        }
        return v.isEmpty() ? null : Path.of(v).toAbsolutePath().normalize();
    }

    private static boolean writeUtf8Append(Path path, String line) {
        try {
            Path dir = path.getParent();
            if (dir != null) {
                Files.createDirectories(dir);
            }
            Files.writeString(
                    path,
                    line,
                    StandardCharsets.UTF_8,
                    StandardOpenOption.CREATE,
                    StandardOpenOption.APPEND);
            return true;
        } catch (Throwable ignored) {
            return false;
        }
    }

    private static Path pathFromEnvOverride() {
        String v = trim(System.getenv("CURSOR_DEBUG_LOG"));
        if (!v.isEmpty()) {
            return Path.of(v).toAbsolutePath().normalize();
        }
        v = trim(System.getenv("PM_AI_DEBUG_LOG"));
        if (!v.isEmpty()) {
            return Path.of(v).toAbsolutePath().normalize();
        }
        return null;
    }

    private static String trim(String s) {
        return s != null ? s.trim() : "";
    }

    /**
     * One NDJSON line with session / hypothesis / data (debug sessions). Does not log secrets; keep {@code data} small.
     */
    public static void appendStructured(
            Map<String, String> ui,
            String sessionId,
            String hypothesisId,
            String location,
            String message,
            Map<String, ?> data) {
        try {
            Map<String, Object> line = new LinkedHashMap<>();
            String sid =
                    sessionId == null || sessionId.isBlank()
                            ? DEFAULT_SESSION_ID
                            : sessionId.trim();
            line.put("sessionId", sid);
            line.put("hypothesisId", hypothesisId);
            line.put("location", location);
            line.put("message", message);
            line.put("data", data != null ? data : Map.of());
            line.put("timestamp", System.currentTimeMillis());
            String json = JSON.writeValueAsString(line);
            appendNdjsonLine(ui, sid, json);
        } catch (Throwable ignored) {
            // debug-only
        }
    }

    /**
     * 配台試行など子プロセス NDJSON と JVM 側 {@link #appendStructured} の {@code sessionId} を揃える。
     *
     * <p>優先: UI の {@code PM_AI_AGENT_DEBUG_SESSION} → OS の同キー → {@code CURSOR_DEBUG_SESSION_ID} →
     * 既定（現在の Cursor デバッグ会話 ID。会話ごとに変える場合は環境変数タブで上書き）。
     */
    public static String resolveDispatchTrialSessionId(Map<String, String> ui) {
        if (ui != null) {
            String s = trim(ui.get("PM_AI_AGENT_DEBUG_SESSION"));
            if (!s.isEmpty()) {
                return s;
            }
        }
        String s2 = trim(System.getenv("PM_AI_AGENT_DEBUG_SESSION"));
        if (!s2.isEmpty()) {
            return s2;
        }
        String s3 = trim(System.getenv("CURSOR_DEBUG_SESSION_ID"));
        if (!s3.isEmpty()) {
            return s3;
        }
        return DEFAULT_SESSION_ID;
    }

    /**
     * Python 子プロセス（段階2・配台試行）向けに {@code PM_AI_AGENT_DEBUG_SESSION} と
     * {@code PM_AI_DEBUG_LOG} を解決して {@code env} に書き込む。環境変数タブで既に非空なら上書きしない。
     *
     * <p>正本: {@code .cursor/rules/agent-debug-ndjson-logging.mdc}
     */
    public static void overlayPythonChildDebugEnv(Map<String, String> env) {
        if (env == null) {
            return;
        }
        String sid = resolveDispatchTrialSessionId(env);
        env.put("PM_AI_AGENT_DEBUG_SESSION", sid);
        String dbg = trim(env.get("PM_AI_DEBUG_LOG"));
        if (dbg.isEmpty()) {
            dbg = trim(env.get(AppPaths.KEY_PM_AI_CURSOR_DEBUG_LOG));
        }
        if (dbg.isEmpty()) {
            env.put(
                    "PM_AI_DEBUG_LOG",
                    resolveNdjsonPath(env, sid).toAbsolutePath().toString());
        }
    }
}
