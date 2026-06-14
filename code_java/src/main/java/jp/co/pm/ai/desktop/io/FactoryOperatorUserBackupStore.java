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
import java.util.UUID;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;

/**
 * {@link AppPaths#FACTORY_OPERATOR_USERS_BIN} の手動バックアップ（世代管理）。
 *
 * <p>保存先: {@link AppPaths#factoryOperatorUsersBackupsRoot} 配下。
 */
public final class FactoryOperatorUserBackupStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final DateTimeFormatter BACKUP_TS =
            DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss").withZone(ZoneId.systemDefault());

    private static final String INDEX_FILE = "index.json";
    private static final String MANIFEST_FILE = "manifest.json";

    /** 手動バックアップの保持上限（世代数）。 */
    public static final int MAX_BACKUP_GENERATIONS = 30;

    public record FactoryOperatorUserBackupEntry(
            String id, String label, long createdAtMillis, String folderName, String createdByOperator) {

        public Path resolveDirectory(Path backupsRoot) {
            String folder = folderName != null && !folderName.isBlank() ? folderName : id;
            return backupsRoot.resolve(folder).toAbsolutePath().normalize();
        }

        public Path resolveBackupFile(Path backupsRoot) {
            return resolveDirectory(backupsRoot).resolve(AppPaths.operatorUsersStoreBinBasename());
        }

        public String displayLabel() {
            if (label != null && !label.isBlank()) {
                return label;
            }
            return id != null ? id : "";
        }
    }

    private FactoryOperatorUserBackupStore() {}

    public static Path resolveBackupsRoot(Map<String, String> ui) {
        return resolveBackupsRoot(ui, null);
    }

    public static Path resolveBackupsRoot(Map<String, String> ui, FactorySite site) {
        String testRoot = System.getProperty("pm.ai.test.factoryOperatorUserBackupRoot");
        if (testRoot != null && !testRoot.isBlank()) {
            return Path.of(testRoot).toAbsolutePath().normalize();
        }
        if (AppPaths.usesRemoteDesktopAppHome()) {
            Map<String, String> u = ui != null ? ui : Map.of();
            return AppPaths.resolveRdpLauncherOperatorUsersBackupsRoot(u);
        }
        Map<String, String> u = ui != null ? ui : Map.of();
        FactorySite effective = site != null ? site : GlobalInitSettingTarget.loadEffective(u);
        return AppPaths.factoryOperatorUsersBackupsRoot(u, effective);
    }

    public static List<FactoryOperatorUserBackupEntry> loadIndex(Map<String, String> ui) {
        Path idx = resolveBackupsRoot(ui).resolve(INDEX_FILE);
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
            List<FactoryOperatorUserBackupEntry> out = new ArrayList<>();
            for (JsonNode n : arr) {
                if (n == null || !n.isObject()) {
                    continue;
                }
                String id = text(n, "id");
                if (id.isBlank()) {
                    continue;
                }
                out.add(
                        new FactoryOperatorUserBackupEntry(
                                id,
                                text(n, "label"),
                                n.path("createdAtMillis").asLong(0L),
                                text(n, "folderName"),
                                text(n, "createdByOperator")));
            }
            out.sort(
                    Comparator.comparingLong(FactoryOperatorUserBackupEntry::createdAtMillis)
                            .reversed());
            return List.copyOf(out);
        } catch (IOException e) {
            return List.of();
        }
    }

    /**
     * schema 昇格前の自動バックアップ（1 プロセス内で同一ストアに対し 1 回まで）。
     *
     * @param priorSchemaVersion 書込前の schemaVersion
     */
    public static void createAutomaticSchemaUpgradeBackup(
            Map<String, String> ui, int priorSchemaVersion, String label) throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        FactorySite effective = FactoryOperatorUserStore.operatorScopeForCurrentApp(u, null);
        FactoryOperatorUserStore.configureForCurrentApp(u, effective);
        Path current = FactoryOperatorUserStore.storePath();
        if (!Files.isRegularFile(current)) {
            return;
        }
        String resolvedLabel =
                label != null && !label.isBlank()
                        ? label.strip()
                        : "アップデート前自動バックアップ schema-" + priorSchemaVersion;
        createManualBackup(u, resolvedLabel);
    }

    /** 現行 {@link FactoryOperatorUserStore} を手動バックアップする。 */
    public static FactoryOperatorUserBackupEntry createManualBackup(Map<String, String> ui, String label)
            throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        FactorySite effective = FactoryOperatorUserStore.operatorScopeForCurrentApp(u, null);
        FactoryOperatorUserStore.configureForCurrentApp(u, effective);
        FactoryOperatorUserStore.ensureStoreFileOnDisk();
        Path current = FactoryOperatorUserStore.storePath();
        if (!Files.isRegularFile(current)) {
            throw new IOException("バックアップ対象のユーザー管理ファイルがありません: " + current);
        }

        Path backupsRoot = resolveBackupsRoot(u, effective);
        Files.createDirectories(backupsRoot);
        String id = UUID.randomUUID().toString().replace("-", "");
        String folder = "backup-" + id;
        Path dir = backupsRoot.resolve(folder);
        Files.createDirectories(dir);
        Files.copy(current, dir.resolve(AppPaths.operatorUsersStoreBinBasename()), StandardCopyOption.REPLACE_EXISTING);

        long now = Instant.now().toEpochMilli();
        String autoLabel =
                BACKUP_TS.format(Instant.ofEpochMilli(now)) + " 手動バックアップ";
        String resolvedLabel = label != null && !label.isBlank() ? label.strip() : autoLabel;
        String operator = FactoryOperatorUserStore.sessionOperatorName();

        ObjectNode manifest = JSON.createObjectNode();
        manifest.put("id", id);
        manifest.put("label", resolvedLabel);
        manifest.put("createdAtMillis", now);
        manifest.put("createdByOperator", operator != null ? operator : "");
        manifest.put("sourcePath", current.toAbsolutePath().normalize().toString());
        JSON.writerWithDefaultPrettyPrinter().writeValue(dir.resolve(MANIFEST_FILE).toFile(), manifest);

        FactoryOperatorUserBackupEntry entry =
                new FactoryOperatorUserBackupEntry(id, resolvedLabel, now, folder, operator);
        List<FactoryOperatorUserBackupEntry> cur = new ArrayList<>(loadIndex(u));
        cur.add(entry);
        trimToMaxGenerations(u, cur, MAX_BACKUP_GENERATIONS);
        saveIndex(u, cur);
        return entry;
    }

    /** 選択したバックアップ世代で現行ファイルを上書き復元する。 */
    public static void restoreFromBackup(FactoryOperatorUserBackupEntry entry, Map<String, String> ui)
            throws IOException {
        if (entry == null) {
            throw new IllegalArgumentException("バックアップが未選択です。");
        }
        Map<String, String> u = ui != null ? ui : Map.of();
        FactorySite effective = GlobalInitSettingTarget.loadEffective(u);
        Path backupsRoot = resolveBackupsRoot(u, effective);
        Path backupFile = entry.resolveBackupFile(backupsRoot);
        if (!Files.isRegularFile(backupFile)) {
            throw new IOException("バックアップファイルが見つかりません: " + backupFile);
        }
        FactoryOperatorUserStore.configureFromUi(u, effective);
        Path target = FactoryOperatorUserStore.storePath();
        if (target.getParent() != null) {
            Files.createDirectories(target.getParent());
        }
        Files.copy(backupFile, target, StandardCopyOption.REPLACE_EXISTING);
    }

    public static void deleteBackupEntry(FactoryOperatorUserBackupEntry entry, Map<String, String> ui)
            throws IOException {
        if (entry == null) {
            return;
        }
        Map<String, String> u = ui != null ? ui : Map.of();
        Path dir = entry.resolveDirectory(resolveBackupsRoot(u));
        if (Files.isDirectory(dir)) {
            deleteDirectoryRecursive(dir);
        }
        List<FactoryOperatorUserBackupEntry> cur = new ArrayList<>(loadIndex(u));
        cur.removeIf(e -> e.id().equals(entry.id()));
        saveIndex(u, cur);
    }

    private static void trimToMaxGenerations(
            Map<String, String> ui, List<FactoryOperatorUserBackupEntry> entries, int max) {
        if (entries.size() <= max) {
            return;
        }
        entries.sort(Comparator.comparingLong(FactoryOperatorUserBackupEntry::createdAtMillis));
        Path backupsRoot = resolveBackupsRoot(ui);
        while (entries.size() > max) {
            FactoryOperatorUserBackupEntry oldest = entries.remove(0);
            try {
                Path dir = oldest.resolveDirectory(backupsRoot);
                if (Files.isDirectory(dir)) {
                    deleteDirectoryRecursive(dir);
                }
            } catch (IOException ignored) {
            }
        }
    }

    private static void saveIndex(Map<String, String> ui, List<FactoryOperatorUserBackupEntry> entries)
            throws IOException {
        Path backupsRoot = resolveBackupsRoot(ui);
        Files.createDirectories(backupsRoot);
        ObjectNode doc = JSON.createObjectNode();
        ArrayNode arr = doc.putArray("entries");
        for (FactoryOperatorUserBackupEntry e : entries) {
            ObjectNode o = arr.addObject();
            o.put("id", e.id());
            o.put("label", e.label() != null ? e.label() : "");
            o.put("createdAtMillis", e.createdAtMillis());
            o.put("folderName", e.folderName() != null ? e.folderName() : "");
            o.put("createdByOperator", e.createdByOperator() != null ? e.createdByOperator() : "");
        }
        JSON.writerWithDefaultPrettyPrinter().writeValue(backupsRoot.resolve(INDEX_FILE).toFile(), doc);
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
