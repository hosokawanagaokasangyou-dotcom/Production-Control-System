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
import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;

/**
 * サマリ {@link AppPaths#SUMMARY_AI_DISPATCH_XLSX} の世代退避（操作者ごと）。
 *
 * <p>最新版は従来どおり {@link AppPaths#summaryAiDispatchXlsxPath} へ上書きする。上書き直前の現行ブックのみ、
 * 操作者別に最大 {@link #MAX_GENERATIONS_PER_USER} 件まで {@link AppPaths#summaryAiDispatchGenerationsRoot}
 * 配下へ退避する。
 */
public final class SummaryAiDispatchGenerationStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final DateTimeFormatter ARCHIVE_TS =
            DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss").withZone(ZoneId.systemDefault());

    private static final String INDEX_FILE = "index.json";
    private static final String MANIFEST_FILE = "manifest.json";
    private static final String WORKBOOK_ARCHIVE_NAME = "workbook.xlsx";
    private static final String UNKNOWN_OPERATOR_DIR = "_unknown";

    /** 操作者あたりの退避上限（ファイル数）。 */
    public static final int MAX_GENERATIONS_PER_USER = 10;

    /** @deprecated {@link #MAX_GENERATIONS_PER_USER} */
    @Deprecated
    public static final int DEFAULT_MAX_GENERATIONS = MAX_GENERATIONS_PER_USER;

    private static final ThreadLocal<Optional<SummaryAiDispatchGenerationEntry>> LAST_ARCHIVED =
            ThreadLocal.withInitial(Optional::empty);

    public record SummaryAiDispatchGenerationEntry(
            String id,
            String operatorUser,
            String label,
            String reason,
            long createdAtMillis,
            String folderName,
            String sourceWorkbookPath) {

        public Path resolveDirectory(Path userGenerationsRoot) {
        String folder = folderName != null && !folderName.isBlank() ? folderName : id;
        return userGenerationsRoot.resolve(folder).toAbsolutePath().normalize();
    }

    public Path resolveDirectoryForOperator(Map<String, String> ui) {
        return resolveDirectory(
                SummaryAiDispatchGenerationStore.resolveOperatorGenerationsRoot(ui, operatorUser));
    }

        public Path resolveWorkbookPath(Path userGenerationsRoot) {
            return resolveDirectory(userGenerationsRoot)
                    .resolve(SummaryAiDispatchGenerationStore.WORKBOOK_ARCHIVE_NAME);
        }

    public Path resolveWorkbookPathForOperator(Map<String, String> ui) {
        return resolveWorkbookPath(
                SummaryAiDispatchGenerationStore.resolveOperatorGenerationsRoot(ui, operatorUser));
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

    /** 操作者別の退避ルート（{@code summary-ai-dispatch-generations/<操作者>/}）。 */
    public static Path resolveUserGenerationsRoot(Map<String, String> ui) {
        return resolveOperatorGenerationsRoot(ui, resolveOperatorUser(ui));
    }

    public static Path resolveOperatorGenerationsRoot(Map<String, String> ui, String operatorUser) {
        String op = operatorUser != null ? operatorUser.strip() : "";
        return resolveGenerationsRoot(ui)
                .resolve(sanitizeOperatorDirName(op.isEmpty() ? UNKNOWN_OPERATOR_DIR : op))
                .toAbsolutePath()
                .normalize();
    }

    public static String resolveOperatorUser(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String fromUi = u.getOrDefault(AppPaths.KEY_PM_AI_OPERATOR_USER, "").strip();
        if (!fromUi.isEmpty()) {
            return fromUi;
        }
        String session = FactoryOperatorUserStore.sessionOperatorName();
        return session.isBlank() ? UNKNOWN_OPERATOR_DIR : session;
    }

    public static boolean isCreatedByCurrentUser(
            SummaryAiDispatchGenerationEntry entry, Map<String, String> ui) {
        if (entry == null) {
            return false;
        }
        String current = resolveOperatorUser(ui);
        String owner = entry.operatorUser() != null ? entry.operatorUser().strip() : "";
        return !owner.isEmpty() && owner.equals(current);
    }

    /** ログイン中操作者の退避一覧（書き込み・上限判定用）。 */
    public static List<SummaryAiDispatchGenerationEntry> loadIndex(Map<String, String> ui) {
        return loadIndexForOperator(ui, resolveOperatorUser(ui));
    }

    /** 全操作者の退避を新しい順に返す（履歴閲覧用）。 */
    public static List<SummaryAiDispatchGenerationEntry> loadAllGenerations(Map<String, String> ui) {
        Path root = resolveGenerationsRoot(ui);
        if (!Files.isDirectory(root)) {
            return List.of();
        }
        List<SummaryAiDispatchGenerationEntry> all = new ArrayList<>();
        try (var stream = Files.list(root)) {
            for (Path sub : stream.filter(Files::isDirectory).sorted().toList()) {
                all.addAll(loadIndexFromUserRoot(sub));
            }
        } catch (IOException ignored) {
            return List.of();
        }
        all.sort(
                Comparator.comparingLong(SummaryAiDispatchGenerationEntry::createdAtMillis)
                        .reversed());
        return List.copyOf(all);
    }

    public static List<SummaryAiDispatchGenerationEntry> loadIndexForOperator(
            Map<String, String> ui, String operatorUser) {
        return loadIndexFromUserRoot(resolveOperatorGenerationsRoot(ui, operatorUser));
    }

    private static List<SummaryAiDispatchGenerationEntry> loadIndexFromUserRoot(Path userRoot) {
        Path idx = userRoot.resolve(INDEX_FILE);
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
            String indexOperator = text(root, "operatorUser");
            List<SummaryAiDispatchGenerationEntry> out = new ArrayList<>();
            for (JsonNode n : arr) {
                if (n == null || !n.isObject()) {
                    continue;
                }
                String id = text(n, "id");
                if (id.isBlank()) {
                    continue;
                }
                String entryOperator = text(n, "operatorUser");
                if (entryOperator.isBlank()) {
                    entryOperator = indexOperator;
                }
                if (entryOperator.isBlank()) {
                    entryOperator = userRoot.getFileName() != null
                            ? userRoot.getFileName().toString()
                            : UNKNOWN_OPERATOR_DIR;
                }
                out.add(
                        new SummaryAiDispatchGenerationEntry(
                                id,
                                entryOperator,
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
        Path archiveWorkbook = entry.resolveWorkbookPathForOperator(ui);
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
        if (!isCreatedByCurrentUser(entry, ui)) {
            throw new IllegalStateException("自分が作成した退避のみ削除できます。");
        }
        String owner = entry.operatorUser() != null ? entry.operatorUser().strip() : resolveOperatorUser(ui);
        Path userRoot = resolveOperatorGenerationsRoot(ui, owner);
        Path dir = entry.resolveDirectory(userRoot);
        if (Files.isDirectory(dir)) {
            deleteDirectoryRecursive(dir);
        }
        List<SummaryAiDispatchGenerationEntry> cur = new ArrayList<>(loadIndexForOperator(ui, owner));
        cur.removeIf(e -> e.id().equals(entry.id()));
        saveIndexForOperator(ui, owner, cur);
    }

    public static void updateEntryLabel(
            SummaryAiDispatchGenerationEntry entry, Map<String, String> ui, String newLabel)
            throws IOException {
        if (entry == null) {
            return;
        }
        if (!isCreatedByCurrentUser(entry, ui)) {
            throw new IllegalStateException("自分が作成した退避のみラベル変更できます。");
        }
        String owner = entry.operatorUser() != null ? entry.operatorUser().strip() : resolveOperatorUser(ui);
        Path userRoot = resolveOperatorGenerationsRoot(ui, owner);
        Path manifest = entry.resolveDirectory(userRoot).resolve(MANIFEST_FILE);
        if (Files.isRegularFile(manifest)) {
            ObjectNode root = (ObjectNode) JSON.readTree(manifest.toFile());
            root.put("label", newLabel != null ? newLabel.strip() : "");
            JSON.writerWithDefaultPrettyPrinter().writeValue(manifest.toFile(), root);
        }
        List<SummaryAiDispatchGenerationEntry> cur = new ArrayList<>(loadIndexForOperator(ui, owner));
        for (int i = 0; i < cur.size(); i++) {
            if (cur.get(i).id().equals(entry.id())) {
                SummaryAiDispatchGenerationEntry old = cur.get(i);
                cur.set(
                        i,
                        new SummaryAiDispatchGenerationEntry(
                                old.id(),
                                old.operatorUser(),
                                newLabel != null ? newLabel.strip() : "",
                                old.reason(),
                                old.createdAtMillis(),
                                old.folderName(),
                                old.sourceWorkbookPath()));
                break;
            }
        }
        saveIndexForOperator(ui, owner, cur);
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

    static String sanitizeOperatorDirName(String operatorUser) {
        if (operatorUser == null || operatorUser.isBlank()) {
            return UNKNOWN_OPERATOR_DIR;
        }
        String t = operatorUser.strip().replaceAll("[\\\\/:*?\"<>|]", "_");
        if (t.length() > 40) {
            t = t.substring(0, 40);
        }
        return t.isEmpty() ? UNKNOWN_OPERATOR_DIR : t;
    }

    private static SummaryAiDispatchGenerationEntry archiveWorkbookCopy(
            Path sourceWorkbook,
            Map<String, String> ui,
            String reason,
            String label,
            String reasonKind)
            throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        String operatorUser = resolveOperatorUser(u);
        Path userRoot = resolveUserGenerationsRoot(u);
        Files.createDirectories(userRoot);
        String id = UUID.randomUUID().toString().replace("-", "");
        String folder = "gen-" + id;
        Path dir = userRoot.resolve(folder);
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
        manifest.put("operatorUser", operatorUser);
        manifest.put("label", resolvedLabel);
        manifest.put("reason", reason != null ? reason.strip() : "");
        manifest.put("reasonKind", reasonKind != null ? reasonKind : "");
        manifest.put("createdAtMillis", now);
        manifest.put("sourceWorkbookPath", sourceWorkbook.toAbsolutePath().normalize().toString());
        JSON.writerWithDefaultPrettyPrinter().writeValue(dir.resolve(MANIFEST_FILE).toFile(), manifest);

        SummaryAiDispatchGenerationEntry entry =
                new SummaryAiDispatchGenerationEntry(
                        id,
                        operatorUser,
                        resolvedLabel,
                        reason != null ? reason.strip() : "",
                        now,
                        folder,
                        sourceWorkbook.toAbsolutePath().normalize().toString());
        List<SummaryAiDispatchGenerationEntry> cur = new ArrayList<>(loadIndex(u));
        cur.add(entry);
        trimToMaxGenerations(u, cur, MAX_GENERATIONS_PER_USER);
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
                Path userRoot = resolveOperatorGenerationsRoot(ui, oldest.operatorUser());
                Path dir = oldest.resolveDirectory(userRoot);
                if (Files.isDirectory(dir)) {
                    deleteDirectoryRecursive(dir);
                }
            } catch (IOException ignored) {
            }
        }
    }

    private static void saveIndex(Map<String, String> ui, List<SummaryAiDispatchGenerationEntry> entries)
            throws IOException {
        saveIndexForOperator(ui, resolveOperatorUser(ui), entries);
    }

    private static void saveIndexForOperator(
            Map<String, String> ui, String operatorUser, List<SummaryAiDispatchGenerationEntry> entries)
            throws IOException {
        Path userRoot = resolveOperatorGenerationsRoot(ui, operatorUser);
        Files.createDirectories(userRoot);
        ObjectNode doc = JSON.createObjectNode();
        doc.put("operatorUser", operatorUser != null ? operatorUser.strip() : "");
        ArrayNode arr = doc.putArray("entries");
        for (SummaryAiDispatchGenerationEntry e : entries) {
            ObjectNode o = arr.addObject();
            o.put("id", e.id());
            o.put("operatorUser", e.operatorUser() != null ? e.operatorUser() : "");
            o.put("label", e.label() != null ? e.label() : "");
            o.put("reason", e.reason() != null ? e.reason() : "");
            o.put("createdAtMillis", e.createdAtMillis());
            o.put("folderName", e.folderName() != null ? e.folderName() : "");
            o.put(
                    "sourceWorkbookPath",
                    e.sourceWorkbookPath() != null ? e.sourceWorkbookPath() : "");
        }
        JSON.writerWithDefaultPrettyPrinter().writeValue(userRoot.resolve(INDEX_FILE).toFile(), doc);
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
