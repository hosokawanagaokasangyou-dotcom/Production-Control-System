package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.UUID;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * 配台ワークスペースのスナップショット（結果 JSON・セッション断片・列順断片）を
 * {@code ~/.pm-ai-desktop/plan-workspace-snapshots/} に保存する。
 */
public final class PlanWorkspaceSnapshotStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final Path ROOT =
            Paths.get(System.getProperty("user.home"), ".pm-ai-desktop", "plan-workspace-snapshots");

    private static final String INDEX_FILE = "index.json";

    public record PlanWorkspaceSnapshotEntry(
            String id,
            String label,
            long createdAtMillis,
            String folderName) {

        public Path resolveDirectory() {
            return ROOT.resolve(folderName != null && !folderName.isBlank() ? folderName : id);
        }
    }

    private PlanWorkspaceSnapshotStore() {}

    /** 工場出荷 UI リセット等で履歴を消す。 */
    public static void deleteAllSilently() {
        try {
            if (!Files.isDirectory(ROOT)) {
                return;
            }
            try (var stream = Files.list(ROOT)) {
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

    public static Path rootDirectory() {
        return ROOT;
    }

    public static List<PlanWorkspaceSnapshotEntry> loadIndex() {
        Path idx = ROOT.resolve(INDEX_FILE);
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
            List<PlanWorkspaceSnapshotEntry> out = new ArrayList<>();
            for (JsonNode n : arr) {
                if (n == null || !n.isObject()) {
                    continue;
                }
                String id = text(n, "id");
                if (id.isBlank()) {
                    continue;
                }
                out.add(
                        new PlanWorkspaceSnapshotEntry(
                                id,
                                text(n, "label"),
                                n.path("createdAtMillis").asLong(0L),
                                text(n, "folderName")));
            }
            out.sort(Comparator.comparingLong(PlanWorkspaceSnapshotEntry::createdAtMillis).reversed());
            return List.copyOf(out);
        } catch (IOException e) {
            return List.of();
        }
    }

    private static void saveIndex(List<PlanWorkspaceSnapshotEntry> entries) throws IOException {
        Files.createDirectories(ROOT);
        ObjectNode root = JSON.createObjectNode();
        ArrayNode arr = root.putArray("entries");
        for (PlanWorkspaceSnapshotEntry e : entries) {
            ObjectNode o = arr.addObject();
            o.put("id", e.id());
            o.put("label", e.label() != null ? e.label() : "");
            o.put("createdAtMillis", e.createdAtMillis());
            o.put("folderName", e.folderName() != null ? e.folderName() : "");
        }
        JSON.writerWithDefaultPrettyPrinter().writeValue(ROOT.resolve(INDEX_FILE).toFile(), root);
    }

    /**
     * 新規スナップショットを作成しインデックスに追加する。
     *
     * @param resultDispatchJsonSource 取り込む {@code 結果_配台表.json} 相当のファイル（必須）
     * @return 作成したエントリ
     */
    public static PlanWorkspaceSnapshotEntry appendSnapshot(
            String label,
            PlanWorkspaceSessionFragment fragment,
            JsonNode columnOrderPartial,
            Path resultDispatchJsonSource)
            throws IOException {
        if (resultDispatchJsonSource == null || !Files.isRegularFile(resultDispatchJsonSource)) {
            throw new IOException("resultDispatchJsonSource が無いかファイルではありません");
        }
        Files.createDirectories(ROOT);
        String id = UUID.randomUUID().toString().replace("-", "");
        String folder = "snap-" + id;
        Path dir = ROOT.resolve(folder);
        Files.createDirectories(dir);

        Files.copy(
                resultDispatchJsonSource,
                dir.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME),
                StandardCopyOption.REPLACE_EXISTING);
        Files.deleteIfExists(dir.resolve("result_dispatch.json"));

        JSON.writerWithDefaultPrettyPrinter()
                .writeValue(dir.resolve("session_fragment.json").toFile(), fragment);

        if (columnOrderPartial != null && columnOrderPartial.isObject() && !columnOrderPartial.isEmpty()) {
            JSON.writerWithDefaultPrettyPrinter()
                    .writeValue(dir.resolve("column_order_partial.json").toFile(), columnOrderPartial);
        } else {
            Files.writeString(dir.resolve("column_order_partial.json"), "{}\n");
        }

        List<PlanWorkspaceSnapshotEntry> cur = new ArrayList<>(loadIndex());
        long now = Instant.now().toEpochMilli();
        String safeLabel = label != null ? label.strip() : "";
        PlanWorkspaceSnapshotEntry entry = new PlanWorkspaceSnapshotEntry(id, safeLabel, now, folder);
        cur.add(entry);
        saveIndex(cur);
        return entry;
    }

    public static void deleteEntry(PlanWorkspaceSnapshotEntry entry) throws IOException {
        if (entry == null) {
            return;
        }
        Path dir = entry.resolveDirectory();
        if (Files.isDirectory(dir)) {
            deleteDirectoryRecursive(dir);
        }
        List<PlanWorkspaceSnapshotEntry> cur = new ArrayList<>(loadIndex());
        cur.removeIf(e -> e.id().equals(entry.id()));
        saveIndex(cur);
    }

    public static PlanWorkspaceSessionFragment readSessionFragment(PlanWorkspaceSnapshotEntry entry)
            throws IOException {
        Path p = entry.resolveDirectory().resolve("session_fragment.json");
        if (!Files.isRegularFile(p)) {
            return PlanWorkspaceSessionFragment.empty();
        }
        return JSON.readValue(p.toFile(), PlanWorkspaceSessionFragment.class);
    }

    public static JsonNode readColumnOrderPartial(PlanWorkspaceSnapshotEntry entry) throws IOException {
        Path p = entry.resolveDirectory().resolve("column_order_partial.json");
        if (!Files.isRegularFile(p)) {
            return JSON.createObjectNode();
        }
        JsonNode n = JSON.readTree(p.toFile());
        return n != null && n.isObject() ? n : JSON.createObjectNode();
    }

    /**
     * スナップショット内の配台表 JSON（段階2と同一ファイル名 {@link AppPaths#RESULT_DISPATCH_TABLE_JSON_BASENAME}。
     * 旧版 {@code result_dispatch.json} のみのフォルダは後方互換で解決する）。
     */
    public static Path resultDispatchJsonPath(PlanWorkspaceSnapshotEntry entry) {
        Path dir = entry.resolveDirectory();
        Path preferred = dir.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME);
        Path legacy = dir.resolve("result_dispatch.json");
        if (Files.isRegularFile(preferred)) {
            return preferred;
        }
        if (Files.isRegularFile(legacy)) {
            return legacy;
        }
        return preferred;
    }

    private static String text(JsonNode n, String key) {
        JsonNode v = n.get(key);
        if (v == null || v.isNull() || !v.isTextual()) {
            return "";
        }
        return v.asText("");
    }

    /** エントリの表示ラベルを更新してインデックスを保存する。 */
    public static void updateEntryLabel(PlanWorkspaceSnapshotEntry entry, String newLabel) throws IOException {
        if (entry == null) {
            return;
        }
        List<PlanWorkspaceSnapshotEntry> cur = new ArrayList<>(loadIndex());
        List<PlanWorkspaceSnapshotEntry> next = new ArrayList<>();
        boolean found = false;
        for (PlanWorkspaceSnapshotEntry e : cur) {
            if (e.id().equals(entry.id())) {
                found = true;
                next.add(
                        new PlanWorkspaceSnapshotEntry(
                                e.id(),
                                newLabel != null ? newLabel.strip() : "",
                                e.createdAtMillis(),
                                e.folderName()));
            } else {
                next.add(e);
            }
        }
        if (found) {
            saveIndex(next);
        }
    }

}
