package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.io.NetworkSourceFileReloadCache;

/**
 * 段階1キャッシュクリア前の退避と、履歴からの復元。
 *
 * <p>保存先: {@code ~/.pm-ai-desktop/workspace-cache-archives/}
 */
public final class WorkspaceCacheArchiveStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final Path DEFAULT_ROOT =
            Paths.get(System.getProperty("user.home"), ".pm-ai-desktop", "workspace-cache-archives");

    private static final String INDEX_FILE = "index.json";
    private static final String MANIFEST_FILE = "manifest.json";

    public record WorkspaceCacheArchiveEntry(
            String id, String label, String reason, long createdAtMillis, String folderName, int fileCount) {

        public Path resolveDirectory() {
            return resolveRoot().resolve(folderName != null && !folderName.isBlank() ? folderName : id);
        }
    }

    public record ArchivedFile(String role, String originalPath, String archiveName) {}

    private WorkspaceCacheArchiveStore() {}

    public static Path rootDirectory() {
        return resolveRoot();
    }

    private static Path resolveRoot() {
        String testRoot = System.getProperty("pm.ai.test.workspaceCacheArchiveRoot");
        if (testRoot != null && !testRoot.isBlank()) {
            return Path.of(testRoot).toAbsolutePath().normalize();
        }
        return DEFAULT_ROOT;
    }

    public static void deleteAllSilently() {
        Path root = resolveRoot();
        try {
            if (!Files.isDirectory(root)) {
                return;
            }
            try (var stream = Files.list(root)) {
                for (Path p : stream.toList()) {
                    if (Files.isDirectory(p)) {
                        deleteDirectoryRecursive(p);
                    } else if (Files.isRegularFile(p)) {
                        Files.deleteIfExists(p);
                    }
                }
            }
        } catch (IOException ignored) {
        }
    }

    private static void deleteDirectoryRecursive(Path dir) throws IOException {
        if (!Files.isDirectory(dir)) {
            return;
        }
        try (var walk = Files.walk(dir)) {
            List<Path> paths = walk.sorted(Comparator.reverseOrder()).toList();
            for (Path p : paths) {
                Files.deleteIfExists(p);
            }
        }
    }

    public static List<WorkspaceCacheArchiveEntry> loadIndex() {
        Path idx = resolveRoot().resolve(INDEX_FILE);
        try {
            if (!Files.isRegularFile(idx)) {
                return List.of();
            }
            JsonNode root = JSON.readTree(idx.toFile());
            if (root == null || !root.isObject()) {
                return List.of();
            }
            JsonNode arr = root.get("entries");
            if (arr == null || !arr.isArray()) {
                return List.of();
            }
            List<WorkspaceCacheArchiveEntry> out = new ArrayList<>();
            for (JsonNode n : arr) {
                if (n == null || !n.isObject()) {
                    continue;
                }
                String id = text(n, "id");
                if (id.isBlank()) {
                    continue;
                }
                out.add(
                        new WorkspaceCacheArchiveEntry(
                                id,
                                text(n, "label"),
                                text(n, "reason"),
                                n.path("createdAtMillis").asLong(0L),
                                text(n, "folderName"),
                                n.path("fileCount").asInt(0)));
            }
            out.sort(
                    Comparator.comparingLong(WorkspaceCacheArchiveEntry::createdAtMillis).reversed());
            return List.copyOf(out);
        } catch (IOException e) {
            return List.of();
        }
    }

    private static void saveIndex(List<WorkspaceCacheArchiveEntry> entries) throws IOException {
        Path storeRoot = resolveRoot();
        Files.createDirectories(storeRoot);
        ObjectNode doc = JSON.createObjectNode();
        ArrayNode arr = doc.putArray("entries");
        for (WorkspaceCacheArchiveEntry e : entries) {
            ObjectNode o = arr.addObject();
            o.put("id", e.id());
            o.put("label", e.label() != null ? e.label() : "");
            o.put("reason", e.reason() != null ? e.reason() : "");
            o.put("createdAtMillis", e.createdAtMillis());
            o.put("folderName", e.folderName() != null ? e.folderName() : "");
            o.put("fileCount", e.fileCount());
        }
        JSON.writerWithDefaultPrettyPrinter()
                .writeValue(storeRoot.resolve(INDEX_FILE).toFile(), doc);
    }

    /**
     * 実在するキャッシュファイルを退避フォルダへコピーする。1 件も無いときは {@code null}。
     */
    public static WorkspaceCacheArchiveEntry archiveDiskCaches(
            Map<String, String> ui, String label, String reason) throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        List<ArchivedFile> archived = new ArrayList<>();
        for (Map.Entry<String, Path> e : Stage1AiCacheClearer.roleToActivePathMap(u).entrySet()) {
            Path src = e.getValue();
            if (!Files.isRegularFile(src)) {
                continue;
            }
            String archiveName = e.getKey() + cacheExtension(src);
            archived.add(new ArchivedFile(e.getKey(), src.toAbsolutePath().normalize().toString(), archiveName));
        }
        if (archived.isEmpty()) {
            return null;
        }

        Path root = resolveRoot();
        Files.createDirectories(root);
        String id = UUID.randomUUID().toString().replace("-", "");
        String folder = "cache-" + id;
        Path dir = root.resolve(folder);
        Files.createDirectories(dir);

        ObjectNode manifest = JSON.createObjectNode();
        manifest.put("reason", reason != null ? reason.strip() : "");
        manifest.put("label", label != null ? label.strip() : "");
        manifest.put("createdAtMillis", Instant.now().toEpochMilli());
        ArrayNode files = manifest.putArray("files");
        for (ArchivedFile f : archived) {
            Files.copy(
                    Path.of(f.originalPath()),
                    dir.resolve(f.archiveName()),
                    StandardCopyOption.REPLACE_EXISTING);
            ObjectNode fo = files.addObject();
            fo.put("role", f.role());
            fo.put("originalPath", f.originalPath());
            fo.put("archiveName", f.archiveName());
        }
        JSON.writerWithDefaultPrettyPrinter().writeValue(dir.resolve(MANIFEST_FILE).toFile(), manifest);

        long now = Instant.now().toEpochMilli();
        WorkspaceCacheArchiveEntry entry =
                new WorkspaceCacheArchiveEntry(
                        id,
                        label != null ? label.strip() : "",
                        reason != null ? reason.strip() : "",
                        now,
                        folder,
                        archived.size());
        List<WorkspaceCacheArchiveEntry> cur = new ArrayList<>(loadIndex());
        cur.add(entry);
        saveIndex(cur);
        return entry;
    }

    public static List<ArchivedFile> readManifestFiles(WorkspaceCacheArchiveEntry entry) throws IOException {
        Path m = entry.resolveDirectory().resolve(MANIFEST_FILE);
        if (!Files.isRegularFile(m)) {
            return List.of();
        }
        JsonNode root = JSON.readTree(m.toFile());
        JsonNode arr = root != null ? root.get("files") : null;
        if (arr == null || !arr.isArray()) {
            return List.of();
        }
        List<ArchivedFile> out = new ArrayList<>();
        for (JsonNode n : arr) {
            if (n == null || !n.isObject()) {
                continue;
            }
            String role = text(n, "role");
            if (role.isBlank()) {
                continue;
            }
            out.add(new ArchivedFile(role, text(n, "originalPath"), text(n, "archiveName")));
        }
        return List.copyOf(out);
    }

    /** 退避内容を現在の解決パスへ復元する。ログ行を返す。 */
    public static List<String> restoreToWorkspace(
            WorkspaceCacheArchiveEntry entry, Map<String, String> ui) throws IOException {
        List<String> logs = new ArrayList<>();
        Map<String, String> u = ui != null ? ui : Map.of();
        Path dir = entry.resolveDirectory();
        Map<String, Path> targets = Stage1AiCacheClearer.roleToActivePathMap(u);
        for (ArchivedFile f : readManifestFiles(entry)) {
            Path archiveFile = dir.resolve(f.archiveName());
            if (!Files.isRegularFile(archiveFile)) {
                logs.add("[cache-archive] 退避ファイルなし: " + f.archiveName());
                continue;
            }
            Path target = targets.get(f.role());
            if (target == null) {
                logs.add("[cache-archive] 復元先不明（role）: " + f.role());
                continue;
            }
            Files.createDirectories(target.getParent());
            Files.copy(archiveFile, target, StandardCopyOption.REPLACE_EXISTING);
            logs.add("[cache-archive] 復元: " + target);
        }
        NetworkSourceFileReloadCache.clearAll();
        logs.add("[cache-archive] メモリ上の再読込キャッシュを破棄しました。");
        return logs;
    }

    public static void deleteEntry(WorkspaceCacheArchiveEntry entry) throws IOException {
        if (entry == null) {
            return;
        }
        Path dir = entry.resolveDirectory();
        if (Files.isDirectory(dir)) {
            deleteDirectoryRecursive(dir);
        }
        List<WorkspaceCacheArchiveEntry> cur = new ArrayList<>(loadIndex());
        cur.removeIf(e -> e.id().equals(entry.id()));
        saveIndex(cur);
    }

    public static void updateEntryLabel(WorkspaceCacheArchiveEntry entry, String newLabel) throws IOException {
        if (entry == null) {
            return;
        }
        Path manifest = entry.resolveDirectory().resolve(MANIFEST_FILE);
        if (Files.isRegularFile(manifest)) {
            ObjectNode root = (ObjectNode) JSON.readTree(manifest.toFile());
            root.put("label", newLabel != null ? newLabel.strip() : "");
            JSON.writerWithDefaultPrettyPrinter().writeValue(manifest.toFile(), root);
        }
        List<WorkspaceCacheArchiveEntry> cur = new ArrayList<>(loadIndex());
        for (int i = 0; i < cur.size(); i++) {
            if (cur.get(i).id().equals(entry.id())) {
                WorkspaceCacheArchiveEntry old = cur.get(i);
                cur.set(
                        i,
                        new WorkspaceCacheArchiveEntry(
                                old.id(),
                                newLabel != null ? newLabel.strip() : "",
                                old.reason(),
                                old.createdAtMillis(),
                                old.folderName(),
                                old.fileCount()));
                break;
            }
        }
        saveIndex(cur);
    }

    private static String cacheExtension(Path src) {
        Path name = src.getFileName();
        if (name == null) {
            return ".bin";
        }
        String s = name.toString();
        int dot = s.lastIndexOf('.');
        return dot >= 0 ? s.substring(dot) : ".bin";
    }

    private static String text(JsonNode n, String field) {
        if (n == null) {
            return "";
        }
        JsonNode v = n.get(field);
        return v != null && !v.isNull() ? v.asText("").strip() : "";
    }
}
