package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.security.MessageDigest;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HexFormat;
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
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;

/**
 * 依頼書入力の受注ファイル Excel 書き込み直前のローカル世代バックアップ。
 *
 * <p>保存先: {@link AppPaths#requestFormJuchuFileBackupsRoot} 配下（受注ファイル絶対パスごとにサブフォルダ）。
 * 同一受注ファイルへのバックアップは {@link #MIN_BACKUP_INTERVAL_MS} 未満の間隔では作成しない。
 */
public final class RequestFormJuchuFileBackupStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final DateTimeFormatter BACKUP_TS =
            DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss").withZone(ZoneId.systemDefault());

    private static final String INDEX_FILE = "index.json";
    private static final String MANIFEST_FILE = "manifest.json";

    /** 同一受注ファイルへの最低バックアップ間隔（15 分）。 */
    public static final long MIN_BACKUP_INTERVAL_MS = 15L * 60L * 1000L;

    /** 受注ファイル（ソースパス）ごとの保持上限（世代数）。超過分は古い順に削除。 */
    public static final int MAX_GENERATIONS_PER_SOURCE = 50;

    /** 一覧ヒント等向け: 退避ポリシー要約（日本語）。 */
    public static String backupPolicySummaryJa() {
        return "書き込み前に自動退避（転記・新規登録等。直近退避から15分未満は省略）";
    }

    /** ボタンツールチップ向け: 退避ポリシー説明（日本語）。 */
    public static String backupPolicyTooltipJa() {
        return "受注ファイルへの書き込み（転記・新規登録・取り消し等）の直前に、"
                + "ローカルへ世代退避したファイルの保存先フォルダを開きます。"
                + "直近の退避から15分未満の連続書き込みでは退避を省略します（15分ごとの定期バックアップではありません）。";
    }

    private static final Object LOCK = new Object();

    private static final ThreadLocal<Optional<RequestFormJuchuFileBackupEntry>> LAST_BACKED_UP =
            ThreadLocal.withInitial(Optional::empty);

    public record RequestFormJuchuFileBackupEntry(
            String id,
            String operatorUser,
            String label,
            String reason,
            long createdAtMillis,
            String folderName,
            String sourceWorkbookPath,
            String archiveFileName) {

        public Path resolveDirectory(Path sourceBackupsRoot) {
            String folder = folderName != null && !folderName.isBlank() ? folderName : id;
            return sourceBackupsRoot.resolve(folder).toAbsolutePath().normalize();
        }

        public Path resolveArchivePath(Path sourceBackupsRoot) {
            String name =
                    archiveFileName != null && !archiveFileName.isBlank()
                            ? archiveFileName
                            : "workbook.xlsm";
            return resolveDirectory(sourceBackupsRoot).resolve(name);
        }

        public String displayLabel() {
            if (label != null && !label.isBlank()) {
                return label;
            }
            return id != null ? id : "";
        }
    }

    private RequestFormJuchuFileBackupStore() {}

    /** 直近 {@link #maybeBackupBeforeWrite} の結果（呼び出しスレッド単位）。 */
    public static Optional<RequestFormJuchuFileBackupEntry> takeLastBackedUp() {
        Optional<RequestFormJuchuFileBackupEntry> last = LAST_BACKED_UP.get();
        LAST_BACKED_UP.remove();
        return last;
    }

    public static Path resolveBackupsRoot(Map<String, String> ui) {
        String testRoot = System.getProperty("pm.ai.test.requestFormJuchuBackupRoot");
        if (testRoot != null && !testRoot.isBlank()) {
            return Path.of(testRoot).toAbsolutePath().normalize();
        }
        return AppPaths.requestFormJuchuFileBackupsRoot(ui);
    }

    /** 受注ファイル絶対パスに対応する世代退避ルート。 */
    public static Path resolveSourceBackupsRoot(Map<String, String> ui, Path sourceWorkbook) {
        Path source = normalizeSource(sourceWorkbook);
        String key = pathKey(source);
        return resolveBackupsRoot(ui).resolve(key).toAbsolutePath().normalize();
    }

    public static List<RequestFormJuchuFileBackupEntry> loadIndexForSource(
            Map<String, String> ui, Path sourceWorkbook) {
        Path idx = resolveSourceBackupsRoot(ui, sourceWorkbook).resolve(INDEX_FILE);
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
            List<RequestFormJuchuFileBackupEntry> out = new ArrayList<>();
            for (JsonNode n : arr) {
                if (n == null || !n.isObject()) {
                    continue;
                }
                String id = text(n, "id");
                if (id.isBlank()) {
                    continue;
                }
                out.add(
                        new RequestFormJuchuFileBackupEntry(
                                id,
                                text(n, "operatorUser"),
                                text(n, "label"),
                                text(n, "reason"),
                                n.path("createdAtMillis").asLong(0L),
                                text(n, "folderName"),
                                text(n, "sourceWorkbookPath"),
                                text(n, "archiveFileName")));
            }
            out.sort(
                    Comparator.comparingLong(RequestFormJuchuFileBackupEntry::createdAtMillis)
                            .reversed());
            return List.copyOf(out);
        } catch (IOException e) {
            return List.of();
        }
    }

    /**
     * 書き込み直前に受注ファイルをローカルへ退避する。直近バックアップから {@link #MIN_BACKUP_INTERVAL_MS}
     * 未満のときはスキップする。
     *
     * @param sourceWorkbook 上書き対象の受注ファイル
     * @param reason 退避理由（例: {@code single-transfer}, {@code bulk-transfer}）
     */
    public static Optional<RequestFormJuchuFileBackupEntry> maybeBackupBeforeWrite(
            Path sourceWorkbook, Map<String, String> ui, String reason) throws IOException {
        Path source = normalizeSource(sourceWorkbook);
        if (source == null || !Files.isRegularFile(source)) {
            LAST_BACKED_UP.set(Optional.empty());
            return Optional.empty();
        }
        synchronized (LOCK) {
            Map<String, String> u = ui != null ? ui : Map.of();
            List<RequestFormJuchuFileBackupEntry> index = loadIndexForSource(u, source);
            if (!index.isEmpty()) {
                long elapsed = Instant.now().toEpochMilli() - index.get(0).createdAtMillis();
                if (elapsed < MIN_BACKUP_INTERVAL_MS) {
                    LAST_BACKED_UP.set(Optional.empty());
                    return Optional.empty();
                }
            }
            RequestFormJuchuFileBackupEntry entry = archiveWorkbookCopy(source, u, reason);
            LAST_BACKED_UP.set(Optional.of(entry));
            return Optional.of(entry);
        }
    }

    public static String reasonLabelJa(String reason) {
        if (reason == null || reason.isBlank()) {
            return "不明";
        }
        return switch (reason) {
            case "single-transfer" -> "単票転記前";
            case "bulk-transfer" -> "一括転記前";
            case "new-order" -> "新規登録前";
            case "undo-transfer" -> "転記取り消し前";
            case "pre-restore" -> "復元前退避";
            default -> reason;
        };
    }

    /**
     * 選択した世代を受注ファイルへ上書き復元する。復元前に現行ファイルがあれば {@code pre-restore} として退避する。
     *
     * @param targetWorkbook 復元先（{@code null} のときは {@code entry} の {@code sourceWorkbookPath}）
     */
    public static void restoreToSourceWorkbook(
            RequestFormJuchuFileBackupEntry entry,
            Map<String, String> ui,
            Path targetWorkbook)
            throws IOException {
        if (entry == null) {
            throw new IllegalArgumentException("バックアップが未選択です。");
        }
        Map<String, String> u = ui != null ? ui : Map.of();
        Path target = resolveRestoreTarget(entry, targetWorkbook);
        if (target == null) {
            throw new IOException("復元先の受注ファイルが未設定です。");
        }
        Path sourceRoot = resolveSourceBackupsRoot(u, target);
        Path archive = entry.resolveArchivePath(sourceRoot);
        if (!Files.isRegularFile(archive)) {
            throw new IOException("バックアップファイルが見つかりません: " + archive);
        }
        synchronized (LOCK) {
            if (Files.isRegularFile(target)) {
                archiveWorkbookCopy(target, u, "pre-restore");
            }
            PoiWorkbookFileWriter.copyFileReplacing(archive, target, u);
        }
    }

    /** 復元先パス（設定中の受注ファイルまたはエントリ記録）。 */
    public static Path resolveRestoreTarget(
            RequestFormJuchuFileBackupEntry entry, Path targetWorkbook) {
        if (targetWorkbook != null && !targetWorkbook.toString().isBlank()) {
            return normalizeSource(targetWorkbook);
        }
        if (entry != null
                && entry.sourceWorkbookPath() != null
                && !entry.sourceWorkbookPath().isBlank()) {
            return normalizeSource(Path.of(entry.sourceWorkbookPath()));
        }
        return null;
    }

    static String pathKey(Path normalizedSource) {
        try {
            MessageDigest md = MessageDigest.getInstance("SHA-256");
            byte[] hash = md.digest(normalizedSource.toString().getBytes(StandardCharsets.UTF_8));
            String hex = HexFormat.of().formatHex(hash, 0, 8);
            String fileName = normalizedSource.getFileName() != null
                    ? normalizedSource.getFileName().toString()
                    : "workbook";
            String safeName = fileName.replaceAll("[\\\\/:*?\"<>|]", "_");
            if (safeName.length() > 40) {
                safeName = safeName.substring(0, 40);
            }
            return hex + "-" + safeName;
        } catch (Exception ex) {
            return UUID.randomUUID().toString().replace("-", "");
        }
    }

    static String resolveOperatorUser(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String fromUi = u.getOrDefault(AppPaths.KEY_PM_AI_OPERATOR_USER, "").strip();
        if (!fromUi.isEmpty()) {
            return fromUi;
        }
        String session = FactoryOperatorUserStore.sessionOperatorName();
        return session.isBlank() ? "_unknown" : session;
    }

    private static Path normalizeSource(Path sourceWorkbook) {
        return sourceWorkbook != null ? sourceWorkbook.toAbsolutePath().normalize() : null;
    }

    private static RequestFormJuchuFileBackupEntry archiveWorkbookCopy(
            Path sourceWorkbook, Map<String, String> ui, String reason) throws IOException {
        String operatorUser = resolveOperatorUser(ui);
        Path sourceRoot = resolveSourceBackupsRoot(ui, sourceWorkbook);
        Files.createDirectories(sourceRoot);

        String id = UUID.randomUUID().toString().replace("-", "");
        String folder = "gen-" + id;
        Path dir = sourceRoot.resolve(folder);
        Files.createDirectories(dir);

        String archiveName =
                sourceWorkbook.getFileName() != null
                        ? sourceWorkbook.getFileName().toString()
                        : "workbook.xlsm";
        Path archivePath = dir.resolve(archiveName);
        Files.copy(sourceWorkbook, archivePath, java.nio.file.StandardCopyOption.REPLACE_EXISTING);

        long now = Instant.now().toEpochMilli();
        String autoLabel =
                BACKUP_TS.format(Instant.ofEpochMilli(now))
                        + " "
                        + sourceWorkbook.getFileName()
                        + " ("
                        + reasonLabelJa(reason)
                        + ")";
        String resolvedReason = reason != null ? reason.strip() : "";

        ObjectNode manifest = JSON.createObjectNode();
        manifest.put("id", id);
        manifest.put("operatorUser", operatorUser);
        manifest.put("label", autoLabel);
        manifest.put("reason", resolvedReason);
        manifest.put("createdAtMillis", now);
        manifest.put("sourceWorkbookPath", sourceWorkbook.toAbsolutePath().normalize().toString());
        manifest.put("archiveFileName", archiveName);
        manifest.put("factorySite", GlobalInitSettingTarget.loadEffective(ui).name());
        JSON.writerWithDefaultPrettyPrinter().writeValue(dir.resolve(MANIFEST_FILE).toFile(), manifest);

        RequestFormJuchuFileBackupEntry entry =
                new RequestFormJuchuFileBackupEntry(
                        id,
                        operatorUser,
                        autoLabel,
                        resolvedReason,
                        now,
                        folder,
                        sourceWorkbook.toAbsolutePath().normalize().toString(),
                        archiveName);

        List<RequestFormJuchuFileBackupEntry> cur =
                new ArrayList<>(loadIndexForSource(ui, sourceWorkbook));
        cur.add(entry);
        trimToMaxGenerations(ui, sourceWorkbook, cur, MAX_GENERATIONS_PER_SOURCE);
        saveIndex(ui, sourceWorkbook, cur);
        return entry;
    }

    private static void trimToMaxGenerations(
            Map<String, String> ui,
            Path sourceWorkbook,
            List<RequestFormJuchuFileBackupEntry> entries,
            int max) {
        if (entries.size() <= max) {
            return;
        }
        entries.sort(Comparator.comparingLong(RequestFormJuchuFileBackupEntry::createdAtMillis));
        Path sourceRoot = resolveSourceBackupsRoot(ui, sourceWorkbook);
        while (entries.size() > max) {
            RequestFormJuchuFileBackupEntry oldest = entries.remove(0);
            try {
                Path dir = oldest.resolveDirectory(sourceRoot);
                if (Files.isDirectory(dir)) {
                    deleteDirectoryRecursive(dir);
                }
            } catch (IOException ignored) {
                // 削除失敗時もインデックス更新は続行
            }
        }
    }

    private static void saveIndex(
            Map<String, String> ui, Path sourceWorkbook, List<RequestFormJuchuFileBackupEntry> entries)
            throws IOException {
        Path sourceRoot = resolveSourceBackupsRoot(ui, sourceWorkbook);
        Files.createDirectories(sourceRoot);
        ObjectNode root = JSON.createObjectNode();
        root.put("sourceWorkbookPath", sourceWorkbook.toAbsolutePath().normalize().toString());
        ArrayNode arr = root.putArray("entries");
        entries.sort(
                Comparator.comparingLong(RequestFormJuchuFileBackupEntry::createdAtMillis)
                        .reversed());
        for (RequestFormJuchuFileBackupEntry e : entries) {
            ObjectNode n = arr.addObject();
            n.put("id", e.id());
            n.put("operatorUser", e.operatorUser());
            n.put("label", e.label());
            n.put("reason", e.reason());
            n.put("createdAtMillis", e.createdAtMillis());
            n.put("folderName", e.folderName());
            n.put("sourceWorkbookPath", e.sourceWorkbookPath());
            n.put("archiveFileName", e.archiveFileName());
        }
        JSON.writerWithDefaultPrettyPrinter().writeValue(sourceRoot.resolve(INDEX_FILE).toFile(), root);
    }

    private static void deleteDirectoryRecursive(Path dir) throws IOException {
        try (var walk = Files.walk(dir)) {
            walk.sorted(Comparator.reverseOrder()).forEach(p -> {
                try {
                    Files.deleteIfExists(p);
                } catch (IOException ignored) {
                    // best effort
                }
            });
        }
    }

    private static String text(JsonNode n, String key) {
        if (n == null || key == null) {
            return "";
        }
        JsonNode v = n.get(key);
        return v != null && !v.isNull() ? v.asText("").strip() : "";
    }
}
