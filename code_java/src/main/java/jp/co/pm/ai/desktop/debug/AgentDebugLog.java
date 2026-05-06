package jp.co.pm.ai.desktop.debug;

import java.lang.management.ManagementFactory;
import java.lang.management.MemoryUsage;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.LinkedHashMap;
import java.util.Map;

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
 */
public final class AgentDebugLog {

    public static final String DEFAULT_SESSION_ID = "e04a1d";

    private static final ObjectMapper JSON = new ObjectMapper();

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
        String uiPath = trim(u.get(AppPaths.KEY_PM_AI_CURSOR_DEBUG_LOG));
        if (!uiPath.isEmpty()) {
            return Path.of(uiPath).toAbsolutePath().normalize();
        }

        Path repo = AppPaths.resolveRepoRoot(u);
        Path parent = repo.getParent();
        if (parent != null) {
            return parent.resolve(".cursor").resolve(fileName).toAbsolutePath().normalize();
        }
        return repo.resolve(".cursor").resolve(fileName).toAbsolutePath().normalize();
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
            return fallback;
        }
        return null;
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
}
