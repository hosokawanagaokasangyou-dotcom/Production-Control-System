package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.UUID;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * サマリ {@link AppPaths#SUMMARY_AI_DISPATCH_XLSX} の世代退避。
 *
 * <p>上書き出力の直前に現行ブックを退避し、{@link AppPaths#summaryAiDispatchGenerationsRoot} 配下の
 * {@code index.json} で履歴を管理する。
 */
public final class SummaryAiDispatchGenerationStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final DateTimeFormatter ARCHIVE_TS =
            DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss").withZone(ZoneId.systemDefault());

    private static final String INDEX_FILE = "index.json";
    private static final String MANIFEST_FILE = "manifest.json";
    private static final String WORKBOOK_ARCHIVE_NAME = "workbook.xlsx";

    public static final int DEFAULT_MAX_GENERATIONS = 30;

    private static final ThreadLocal<Optional<SummaryAiDispatchGenerationEntry>> LAST_ARCHIVED =
            ThreadLocal.withInitial(Optional::empty);

    public record SummaryAiDispatchGenerationEntry(
            String id,
            String label,
            String reason,
            long createdAtMillis,
            String folderName,
            String sourceWorkbookPath) {

        public Path resolveDirectory(Path generationsRoot) {
            String folder = folderName != null && !folderName.isBlank() ? folderName : id;
            return generationsRoot.resolve(folder).toAbsolutePath().normalize();
        }

        public Path resolveWorkbookPath(Path generationsRoot) {
            return resolveDirectory(generationsRoot).resolve(WORKBOOK_ARCHIVE_NAME);
        }

        public String displayLabel() {
            if (label != null && !label.isBlank()) {
                return label;
            }
            return id != null ? id : "";
        }
    }

    private SummaryAiDispatchGenerationStore() {}

    /** 直近 {@link #archiveBeforeOverwrite} / {@link #archiveCurrent} の結果（出力スレッド単位）。 */
    public static Optional<SummaryAiDispatchGenerationEntry> takeLastArchived() {
        Optional<SummaryAiDispatchGenerationEntry> last = LAST_ARCHIVED.get();
        LAST_ARCHIVED.remove();
        return last;
    }

    public static Path resolveGenerationsRoot(Map<String, String> ui) {
        String testRoot = System.getProperty("pm.ai.test.summaryGenerationRoot");
        if (testRoot != null && !testRoot.isBlank()) {
            return Path.of(testRoot).toAbsolutePath().normalize();
        }
        return AppPaths.summaryAiDispatchGenerationsRoot(ui);
    }

    public static List<SummaryAiDispatchGenerationEntry> loadIndex(Map<String, String> ui) {
        Path idx = resolveGenerationsRoot(ui).resolve(INDEX_FILE);
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
            List<SummaryAiDispatchGenerationEntry> out = new ArrayList<>();
            for (JsonNode n : arr) {
                if (n == null || !n.isObject()) {
                    continue;
                }
                String id = text(n, "id");
                if (id.isBlank()) {
                    continue;
                }
                out.add(
                        new SummaryAiDispatchGenerationEntry(
                                id,
                                text(n, "label"),
                                text(n, "reason"),
                                n.path("createdAtMillis").asLong(0L),
                                text(n, "folderName"),
                                text(n, "sourceWorkbookPath")));
            }
            out.sort(
                    Comparator.comparingLong(SummaryAiDispatchGenerationEntry::createdAtMillis)
                            .reversed());
            return List.copyOf(out);
        } catch (IOException e) {
            return List.of();
        }
    }

    /**
     * 上書き直前に現行ブックを退避する。対象ファイルが無いときは空。
     *
     * @param currentWorkbook {@link AppPaths#summaryAiDispatchXlsxPath} 等の上書き対象
     * @param reason 退避理由（例: {@code delivery-reload}, {@code stage3-export}）
     */
    public static Optional<SummaryAiDispatchGenerationEntry> archiveBeforeOverwrite(
            Path currentWorkbook, Map<String, String> ui, String reason) throws IOException {
        Path current = currentWorkbook != null ? currentWorkbook.toAbsolutePath().normalize() : null;
        if (current == null || !Files.isRegularFile(current)) {
            LAST_ARCHIVED.set(Optional.empty());
            return Optional.empty();
        }
        SummaryAiDispatchGenerationEntry entry =
                archiveWorkbookCopy(current, ui, reason, null, "overwrite");
        LAST_ARCHIVED.set(Optional.of(entry));
        return Optional.of(entry);
    }

    /** 手動退避（現行ブックをコピーするだけ。上書きはしない）。 */
    public static Optional<SummaryAiDispatchGenerationEntry> archiveCurrent(
            Map<String, String> ui, String label) throws IOException {
        Path current = AppPaths.summaryAiDispatchXlsxPath(ui);
        if (!Files.isRegularFile(current)) {
            return Optional.empty();
        }
        SummaryAiDispatchGenerationEntry entry =
                archiveWorkbookCopy(current, ui, "manual-archive", label, "manual");
        LAST_ARCHIVED.set(Optional.of(entry));
        return Optional.of(entry);
    }

    /** 選択した世代を現行ブックへ復元する（復元前に現行を退避）。 */
    public static void restoreToCurrentWorkbook(
            SummaryAiDispatchGenerationEntry entry, Map<String, String> ui) throws IOException {
        if (entry == null) {
            return;
        }
        Path generationsRoot = resolveGenerationsRoot(ui);
        Path archiveWorkbook = entry.resolveWorkbookPath(generationsRoot);
        if (!Files.isRegularFile(archiveWorkbook)) {
            throw new IOException("退避ブックが見つかりません: " + archiveWorkbook);
        }
        Path current = AppPaths.summaryAiDispatchXlsxPath(ui);
        archiveBeforeOverwrite(current, ui, "restore-backup");
        if (current.getParent() != null) {
            Files.createDirectories(current.getParent());
        }
        Files.copy(archiveWorkbook, current, StandardCopyOption.REPLACE_EXISTING);
    }

    public static void deleteEntry(SummaryAiDispatchGenerationEntry entry, Map<String, String> ui)
            throws IOException {
        if (entry == null) {
            return;
        }
        Path dir = entry.resolveDirectory(resolveGenerationsRoot(ui));
        if (Files.isDirectory(dir)) {
            deleteDirectoryRecursive(dir);
        }
        List<SummaryAiDispatchGenerationEntry> cur = new ArrayList<>(loadIndex(ui));
        cur.removeIf(e -> e.id().equals(entry.id()));
        saveIndex(ui, cur);
    }

    public static void updateEntryLabel(
            SummaryAiDispatchGenerationEntry entry, Map<String, String> ui, String newLabel)
            throws IOException {
        if (entry == null) {
            return;
        }
        Path manifest = entry.resolveDirectory(resolveGenerationsRoot(ui)).resolve(MANIFEST_FILE);
        if (Files.isRegularFile(manifest)) {
            ObjectNode root = (ObjectNode) JSON.readTree(manifest.toFile());
            root.put("label", newLabel != null ? newLabel.strip() : "");
            JSON.writerWithDefaultPrettyPrinter().writeValue(manifest.toFile(), root);
        }
        List<SummaryAiDispatchGenerationEntry> cur = new ArrayList<>(loadIndex(ui));
        for (int i = 0; i < cur.size(); i++) {
            if (cur.get(i).id().equals(entry.id())) {
                SummaryAiDispatchGenerationEntry old = cur.get(i);
                cur.set(
                        i,
                        new SummaryAiDispatchGenerationEntry(
                                old.id(),
                                newLabel != null ? newLabel.strip() : "",
                                old.reason(),
                                old.createdAtMillis(),
                                old.folderName(),
                                old.sourceWorkbookPath()));
                break;
            }
        }
        saveIndex(ui, cur);
    }

    public static String reasonLabelJa(String reason) {
        if (reason == null || reason.isBlank()) {
            return "不明";
        }
        return switch (reason) {
            case "overwrite" -> "上書き前";
            case "manual" -> "手動退避";
            case "manual-archive" -> "手動退避";
            case "delivery-reload" -> "納期再読込出力";
            case "stage3-export" -> "段階3後出力";
            case "pipeline-export" -> "パイプライン出力";
            case "export" -> "サマリ出力";
            case "restore-backup" -> "復元前退避";
            default -> reason;
        };
    }

    private static SummaryAiDispatchGenerationEntry archiveWorkbookCopy(
            Path sourceWorkbook,
            Map<String, String> ui,
            String reason,
            String label,
            String reasonKind)
            throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path generationsRoot = resolveGenerationsRoot(u);
        Files.createDirectories(generationsRoot);
        String id = UUID.randomUUID().toString().replace("-", "");
        String folder = "gen-" + id;
        Path dir = generationsRoot.resolve(folder);
        Files.createDirectories(dir);
        Files.copy(
                sourceWorkbook,
                dir.resolve(WORKBOOK_ARCHIVE_NAME),
                StandardCopyOption.REPLACE_EXISTING);

        long now = Instant.now().toEpochMilli();
        String autoLabel =
                ARCHIVE_TS.format(Instant.ofEpochMilli(now))
                        + " "
                        + sourceWorkbook.getFileName();
        String resolvedLabel =
                label != null && !label.isBlank() ? label.strip() : autoLabel;

        ObjectNode manifest = JSON.createObjectNode();
        manifest.put("id", id);
        manifest.put("label", resolvedLabel);
        manifest.put("reason", reason != null ? reason.strip() : "");
        manifest.put("reasonKind", reasonKind != null ? reasonKind : "");
        manifest.put("createdAtMillis", now);
        manifest.put("sourceWorkbookPath", sourceWorkbook.toAbsolutePath().normalize().toString());
        JSON.writerWithDefaultPrettyPrinter().writeValue(dir.resolve(MANIFEST_FILE).toFile(), manifest);

        SummaryAiDispatchGenerationEntry entry =
                new SummaryAiDispatchGenerationEntry(
                        id,
                        resolvedLabel,
                        reason != null ? reason.strip() : "",
                        now,
                        folder,
                        sourceWorkbook.toAbsolutePath().normalize().toString());
        List<SummaryAiDispatchGenerationEntry> cur = new ArrayList<>(loadIndex(u));
        cur.add(entry);
        trimToMaxGenerations(u, cur, DEFAULT_MAX_GENERATIONS);
        saveIndex(u, cur);
        return entry;
    }

    private static void trimToMaxGenerations(
            Map<String, String> ui, List<SummaryAiDispatchGenerationEntry> entries, int max) {
        if (entries.size() <= max) {
            return;
        }
        entries.sort(
                Comparator.comparingLong(SummaryAiDispatchGenerationEntry::createdAtMillis));
        while (entries.size() > max) {
            SummaryAiDispatchGenerationEntry oldest = entries.remove(0);
            try {
                Path dir = oldest.resolveDirectory(resolveGenerationsRoot(ui));
                if (Files.isDirectory(dir)) {
                    deleteDirectoryRecursive(dir);
                }
            } catch (IOException ignored) {
            }
        }
    }

    private static void saveIndex(Map<String, String> ui, List<SummaryAiDispatchGenerationEntry> entries)
            throws IOException {
        Path generationsRoot = resolveGenerationsRoot(ui);
        Files.createDirectories(generationsRoot);
        ObjectNode doc = JSON.createObjectNode();
        ArrayNode arr = doc.putArray("entries");
        for (SummaryAiDispatchGenerationEntry e : entries) {
            ObjectNode o = arr.addObject();
            o.put("id", e.id());
            o.put("label", e.label() != null ? e.label() : "");
            o.put("reason", e.reason() != null ? e.reason() : "");
            o.put("createdAtMillis", e.createdAtMillis());
            o.put("folderName", e.folderName() != null ? e.folderName() : "");
            o.put(
                    "sourceWorkbookPath",
                    e.sourceWorkbookPath() != null ? e.sourceWorkbookPath() : "");
        }
        JSON.writerWithDefaultPrettyPrinter()
                .writeValue(generationsRoot.resolve(INDEX_FILE).toFile(), doc);
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

    private static String text(JsonNode n, String field) {
        if (n == null) {
            return "";
        }
        JsonNode v = n.get(field);
        return v != null && !v.isNull() ? v.asText("").strip() : "";
    }
}
