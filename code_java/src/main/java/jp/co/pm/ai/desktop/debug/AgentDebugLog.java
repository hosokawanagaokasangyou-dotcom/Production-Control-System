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
import java.util.Map;

import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * NDJSON append for Cursor debug sessions. Resolves log path across Windows / WSL / fat JAR layouts.
 *
 * <p>Project convention: {@code .cursor/rules/agent-debug-ndjson-logging.mdc}（モノレポでは親
 * {@code 工程管理AIプロジェクト_JAVA/.cursor/rules/agent-debug-ndjson-logging.mdc} が正本）。
 *
 * <p>書き込みは複数候補を順に試す（いずれかが成功すればよい）。候補の組み立ては {@link #resolveNdjsonPath} と整合。
 *
 * <p>概略（重複パスは除外）:
 *
 * <ol>
 *   <li>{@code CURSOR_DEBUG_LOG} / {@code PM_AI_DEBUG_LOG}</li>
 *   <li>UI {@link AppPaths#KEY_PM_AI_CURSOR_DEBUG_LOG}</li>
 *   <li>{@code parent(resolveRepoRoot) / .cursor / …}（リーフが {@code Production-Control-System} のとき）</li>
 *   <li>{@code resolveRepoRoot / .cursor / …}</li>
 *   <li>{@code user.dir} を遡り、PM AI リポジトリらしき根に {@code .cursor / …}</li>
 *   <li>{@code user.dir} を遡り、最初に見つかった {@code .cursor / …}</li>
 *   <li>{@code user.home / .cursor / …}</li>
 *   <li>{@code java.io.tmpdir / pm-ai-agent-debug-&lt;sessionId&gt;.ndjson}</li>
 * </ol>
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
        List<Path> c = ndjsonWriteCandidates(ui, sessionId);
        return c.isEmpty() ? tmpdirNdjsonPath(sessionId) : c.get(0);
    }

    /**
     * 実際に追記を試みるパス候補（先頭ほど優先）。{@link #appendNdjsonLine} と同一順。
     */
    public static List<Path> ndjsonWriteCandidates(Map<String, String> ui, String sessionId) {
        String id =
                sessionId == null || sessionId.isBlank()
                        ? DEFAULT_SESSION_ID
                        : sessionId.trim();
        String fileName = "debug-" + id + ".log";

        LinkedHashSet<Path> seen = new LinkedHashSet<>();
        List<Path> out = new ArrayList<>();

        Path envPath = pathFromEnvOverride();
        if (envPath != null) {
            addCandidate(out, seen, envPath);
            return out;
        }

        Map<String, String> u = ui != null ? ui : Map.of();
        String uiPath = trim(u.get(AppPaths.KEY_PM_AI_CURSOR_DEBUG_LOG));
        if (!uiPath.isEmpty()) {
            addCandidate(out, seen, Path.of(uiPath).toAbsolutePath().normalize());
            return out;
        }

        Path repo = AppPaths.resolveRepoRoot(u);
        Path parent = repo.getParent();
        if (parent != null && isProductionControlSystemRepoLeaf(repo)) {
            addCandidate(
                    out,
                    seen,
                    parent.resolve(".cursor").resolve(fileName).toAbsolutePath().normalize());
        }
        addCandidate(out, seen, repo.resolve(".cursor").resolve(fileName).toAbsolutePath().normalize());

        Path walkedRepo = walkUpPmAiRepoCursorLog(fileName);
        if (walkedRepo != null) {
            addCandidate(out, seen, walkedRepo);
        }
        Path walkedCursor = walkUpFirstExistingCursorLog(fileName);
        if (walkedCursor != null) {
            addCandidate(out, seen, walkedCursor);
        }

        Path homeFallback =
                Path.of(System.getProperty("user.home", "."))
                        .resolve(".cursor")
                        .resolve(fileName)
                        .toAbsolutePath()
                        .normalize();
        addCandidate(out, seen, homeFallback);

        addCandidate(out, seen, tmpdirNdjsonPath(sessionId));
        return out;
    }

    private static void addCandidate(List<Path> out, LinkedHashSet<Path> seen, Path p) {
        if (p == null) {
            return;
        }
        Path n = p.toAbsolutePath().normalize();
        if (seen.add(n)) {
            out.add(n);
        }
    }

    /**
     * {@code user.dir} から上位へ遡り、リポジトリらしきディレクトリ直下の {@code .cursor/debug-…​.log}。
     */
    private static Path walkUpPmAiRepoCursorLog(String fileName) {
        Path start = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
        Path p = start;
        for (int depth = 0; depth < 14 && p != null; depth++) {
            if (looksLikePmAiRepositoryRoot(p)) {
                return p.resolve(".cursor").resolve(fileName).toAbsolutePath().normalize();
            }
            p = p.getParent();
        }
        return null;
    }

    /**
     * {@code user.dir} から上位へ遡り、最初に存在する {@code .cursor} ディレクトリ配下のログパス。
     */
    private static Path walkUpFirstExistingCursorLog(String fileName) {
        Path start = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
        Path p = start;
        for (int depth = 0; depth < 14 && p != null; depth++) {
            Path cursorDir = p.resolve(".cursor");
            if (Files.isDirectory(cursorDir)) {
                return cursorDir.resolve(fileName).toAbsolutePath().normalize();
            }
            p = p.getParent();
        }
        return null;
    }

    private static boolean looksLikePmAiRepositoryRoot(Path p) {
        if (p == null || !Files.isDirectory(p)) {
            return false;
        }
        try {
            if (Files.isRegularFile(p.resolve("version.txt"))) {
                return true;
            }
            Path pom = p.resolve("code_java").resolve("pom.xml");
            if (Files.isRegularFile(pom)) {
                return true;
            }
            Path py = p.resolve("code").resolve("python");
            return Files.isDirectory(py);
        } catch (Throwable ignored) {
            return false;
        }
    }

    private static Path tmpdirNdjsonPath(String sessionId) {
        String id =
                sessionId == null || sessionId.isBlank()
                        ? DEFAULT_SESSION_ID
                        : sessionId.trim();
        return Path.of(System.getProperty("java.io.tmpdir", "."))
                .resolve("pm-ai-agent-debug-" + id + ".ndjson")
                .toAbsolutePath()
                .normalize();
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
        List<Path> candidates = ndjsonWriteCandidates(ui, sessionId);
        for (Path path : candidates) {
            if (writeUtf8Append(path, line)) {
                return path;
            }
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
