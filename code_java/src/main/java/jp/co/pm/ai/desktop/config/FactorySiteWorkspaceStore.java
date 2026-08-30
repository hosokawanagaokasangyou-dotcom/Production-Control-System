package jp.co.pm.ai.desktop.config;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.ScheduledFuture;
import java.util.concurrent.TimeUnit;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * 操作者×工場のワークスペース（uiEnvRows + 工場スコープ session）を {@code ~/.pm-ai-desktop/operator-local/} に保存。
 */
public final class FactorySiteWorkspaceStore {

    public static final int SCHEMA_VERSION = 1;

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final ConcurrentHashMap<CacheKey, FactorySiteWorkspaceSnapshot> MEMORY =
            new ConcurrentHashMap<>();

    private static final ScheduledExecutorService FLUSH_EXECUTOR =
            Executors.newSingleThreadScheduledExecutor(
                    r -> {
                        Thread t = new Thread(r, "factory-workspace-flush");
                        t.setDaemon(true);
                        return t;
                    });

    private static volatile ScheduledFuture<?> pendingFlush;

    private record CacheKey(String operatorSlug, FactorySite site) {}

    private FactorySiteWorkspaceStore() {}

    public static void save(String operatorName, FactorySite site, FactorySiteWorkspaceSnapshot snapshot) {
        Optional<String> slug = FactoryOperatorUserStore.operatorLocalStorageSlug(operatorName);
        if (slug.isEmpty() || site == null || site == FactorySite.RDP_LAUNCHER || snapshot == null) {
            return;
        }
        MEMORY.put(new CacheKey(slug.get(), site), snapshot);
        scheduleDebouncedFlush(slug.get());
    }

    public static Optional<FactorySiteWorkspaceSnapshot> load(String operatorName, FactorySite site) {
        Optional<String> slug = FactoryOperatorUserStore.operatorLocalStorageSlug(operatorName);
        if (slug.isEmpty() || site == null || site == FactorySite.RDP_LAUNCHER) {
            return Optional.empty();
        }
        CacheKey key = new CacheKey(slug.get(), site);
        FactorySiteWorkspaceSnapshot mem = MEMORY.get(key);
        if (mem != null) {
            return Optional.of(mem);
        }
        Optional<FactorySiteWorkspaceSnapshot> disk = readFromDisk(operatorName, site);
        disk.ifPresent(s -> MEMORY.put(key, s));
        return disk;
    }

    /** 対象工場に保存された、現在も一覧可能な依頼書原本フォルダを返す。 */
    public static Optional<String> loadReachableRequestFormOriginalDir(
            String operatorName, FactorySite site) {
        return load(operatorName, site)
                .flatMap(
                        snapshot ->
                                snapshot.uiEnvRows().stream()
                                        .filter(
                                                row ->
                                                        AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR
                                                                .equals(row.name()))
                                        .map(UiEnvRowSnapshot::value)
                                        .map(value -> value != null ? value.strip() : "")
                                        .filter(value -> !value.isEmpty())
                                        .map(
                                                value -> {
                                                    try {
                                                        return Path.of(value)
                                                                .toAbsolutePath()
                                                                .normalize();
                                                    } catch (RuntimeException ex) {
                                                        return null;
                                                    }
                                                })
                                        .filter(NetworkSourceDirResolver::isDirectoryListingReachable)
                                        .map(Path::toString)
                                        .findFirst());
    }

    public static void saveLastFactorySite(String operatorName, FactorySite site) {
        if (FactoryOperatorUserStore.operatorLocalStorageSlug(operatorName).isEmpty()
                || site == null
                || site == FactorySite.RDP_LAUNCHER) {
            return;
        }
        try {
            Files.createDirectories(AppPaths.operatorLocalWorkspaceDir(operatorName).orElseThrow());
            Files.writeString(
                    AppPaths.operatorLastFactorySitePath(operatorName),
                    site.name(),
                    StandardCharsets.UTF_8);
        } catch (Exception ignored) {
        }
    }

    public static Optional<FactorySite> loadLastFactorySite(String operatorName) {
        if (FactoryOperatorUserStore.operatorLocalStorageSlug(operatorName).isEmpty()) {
            return Optional.empty();
        }
        try {
            Path p = AppPaths.operatorLastFactorySitePath(operatorName);
            if (!Files.isRegularFile(p)) {
                return Optional.empty();
            }
            String raw = Files.readString(p, StandardCharsets.UTF_8).trim();
            if (raw.isEmpty()) {
                return Optional.empty();
            }
            return Optional.of(FactorySite.valueOf(raw));
        } catch (Exception ignored) {
            return Optional.empty();
        }
    }

    public static void warmMemoryCacheFromDisk(String operatorName) {
        for (FactorySite site : FactorySite.dispatchProductionSites()) {
            load(operatorName, site);
        }
    }

    public static void flushMemoryCacheToDisk(String operatorName) {
        Optional<String> slug = FactoryOperatorUserStore.operatorLocalStorageSlug(operatorName);
        if (slug.isEmpty()) {
            return;
        }
        String s = slug.get();
        for (FactorySite site : FactorySite.dispatchProductionSites()) {
            CacheKey key = new CacheKey(s, site);
            FactorySiteWorkspaceSnapshot snap = MEMORY.get(key);
            if (snap != null) {
                writeToDisk(operatorName, site, snap);
            }
        }
    }

    public static void onOperatorSessionChanged(String oldOperator, String newOperator) {
        if (oldOperator != null && !oldOperator.isBlank()) {
            flushMemoryCacheToDisk(oldOperator);
        }
        if (newOperator != null && !newOperator.isBlank()) {
            warmMemoryCacheFromDisk(newOperator);
        }
    }

    public static Path pathFor(String operatorName, FactorySite site) {
        return AppPaths.operatorFactoryWorkspacePath(operatorName, site);
    }

    /** テスト用: メモリキャッシュと pending flush をクリア。 */
    public static void resetForTests() {
        MEMORY.clear();
        ScheduledFuture<?> f = pendingFlush;
        if (f != null) {
            f.cancel(false);
        }
        pendingFlush = null;
    }

    private static void scheduleDebouncedFlush(String operatorSlug) {
        ScheduledFuture<?> prev = pendingFlush;
        if (prev != null) {
            prev.cancel(false);
        }
        pendingFlush =
                FLUSH_EXECUTOR.schedule(
                        () -> flushAllSlugsMatching(operatorSlug), 500, TimeUnit.MILLISECONDS);
    }

    private static void flushAllSlugsMatching(String operatorSlug) {
        String operatorName = FactoryOperatorUserStore.sessionOperatorName();
        String sessionSlug =
                FactoryOperatorUserStore.operatorLocalStorageSlug(operatorName).orElse("");
        String nameForPath = sessionSlug.equals(operatorSlug) ? operatorName : operatorSlug;
        for (var e : MEMORY.entrySet()) {
            if (e.getKey().operatorSlug().equals(operatorSlug)) {
                writeToDisk(nameForPath, e.getKey().site(), e.getValue());
            }
        }
    }

    private static Optional<FactorySiteWorkspaceSnapshot> readFromDisk(
            String operatorName, FactorySite site) {
        try {
            Path path = AppPaths.operatorFactoryWorkspacePath(operatorName, site);
            if (!Files.isRegularFile(path)) {
                return Optional.empty();
            }
            JsonNode root = JSON.readTree(path.toFile());
            if (root == null || !root.isObject()) {
                return Optional.empty();
            }
            List<UiEnvRowSnapshot> rows = parseUiEnvRows(root.get("uiEnvRows"));
            DesktopSessionState session =
                    DesktopSessionStateStore.parseSessionFragment(root.get("session"));
            return Optional.of(new FactorySiteWorkspaceSnapshot(rows, session));
        } catch (Exception ignored) {
            return Optional.empty();
        }
    }

    private static void writeToDisk(
            String operatorName, FactorySite site, FactorySiteWorkspaceSnapshot snapshot) {
        try {
            Path path = AppPaths.operatorFactoryWorkspacePath(operatorName, site);
            Files.createDirectories(path.getParent());
            ObjectNode root = JSON.createObjectNode();
            root.put("schemaVersion", SCHEMA_VERSION);
            ArrayNode arr = root.putArray("uiEnvRows");
            for (UiEnvRowSnapshot row : snapshot.uiEnvRows()) {
                ObjectNode o = arr.addObject();
                o.put("name", row.name() != null ? row.name() : "");
                o.put("value", row.value() != null ? row.value() : "");
                o.put("description", row.description() != null ? row.description() : "");
            }
            root.set("session", DesktopSessionStateStore.toJsonObject(snapshot.sessionFragment()));
            JSON.writeValue(path.toFile(), root);
        } catch (Exception ignored) {
        }
    }

    private static List<UiEnvRowSnapshot> parseUiEnvRows(JsonNode arr) {
        if (arr == null || !arr.isArray()) {
            return List.of();
        }
        List<UiEnvRowSnapshot> out = new ArrayList<>();
        for (JsonNode n : arr) {
            if (n == null || !n.isObject()) {
                continue;
            }
            out.add(
                    new UiEnvRowSnapshot(
                            n.path("name").asText(""),
                            n.path("value").asText(""),
                            n.path("description").asText("")));
        }
        return List.copyOf(out);
    }
}
