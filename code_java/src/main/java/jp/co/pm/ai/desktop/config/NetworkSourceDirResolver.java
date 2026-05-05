package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Stream;

/**
 * {@link AppPaths#KEY_PM_AI_TASK_INPUT_SOURCE_DIR} / {@link AppPaths#KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR}
 * 由来のファイルがネットワーク等で参照できないとき、リポジトリ配下キャッシュの最終成功コピーを Python 子プロセス向け
 * 環境変数（{@code PM_AI_PROCESSING_PLAN_PATH} / {@code PM_AI_ACTUAL_DETAIL_WORKBOOK}）へフォールバックする。
 *
 * <p>優先順位は planning_core の {@code dispatch_workspace.resolve_processing_plan_path_from_env} および
 * {@code resolve_actual_detail_workbook_path} に概ね合わせる。
 */
public final class NetworkSourceDirResolver {

    private static final String META_JSON = "network-source-cache-meta.json";

    private static final List<String> TASK_INPUT_SUFFIXES =
            List.of(".csv", ".parquet", ".pq", ".xlsx", ".xlsm", ".xltx", ".xltm");

    /**
     * @param taskInputFromCache {@code true} iff effective task-input file is read from local cache
     * @param actualDetailFromCache {@code true} iff effective actual-detail workbook is read from local cache
     */
    public record Result(
            Optional<Path> taskInputPath,
            boolean taskInputFromCache,
            Optional<Path> actualDetailPath,
            boolean actualDetailFromCache,
            List<String> logLines) {}

    private NetworkSourceDirResolver() {}

    /**
     * 環境マップ {@code m} から加工計画ファイル・実績明細ブックを解決する。
     *
     * @param skipTaskInputSourceDirListing {@code true} のとき {@link AppPaths#resolveTaskInputSourceDir(Map)}
     *     配下の一覧・最新ファイル検出をせず、単一ファイル指定が無効な場合はキャッシュのみ試行する（起動時未到達など）。
     * @param skipActualDetailSourceDirListing 同上 {@link AppPaths#resolveActualDetailSourceDir(Map)}
     */
    public static Result resolve(
            Map<String, String> m,
            boolean skipTaskInputSourceDirListing,
            boolean skipActualDetailSourceDirListing) {
        List<String> logs = new ArrayList<>();
        Optional<Path> task = resolveTaskInput(m, logs, skipTaskInputSourceDirListing);
        Optional<Path> actual = resolveActualDetail(m, logs, skipActualDetailSourceDirListing);
        boolean tCache =
                task.isPresent()
                        && task.get().startsWith(cacheRoot(m))
                        && Files.isRegularFile(task.get());
        boolean aCache =
                actual.isPresent()
                        && actual.get().startsWith(cacheRoot(m))
                        && Files.isRegularFile(actual.get());
        return new Result(task, tCache, actual, aCache, List.copyOf(logs));
    }

    /** フォルダ一覧まで試す通常解決（後方互換）。 */
    public static Result resolve(Map<String, String> m) {
        return resolve(m, false, false);
    }

    /**
     * 環境変数で解決されるソースフォルダにディレクトリとしてアクセスできるか（一覧が開けるか）。
     * 起動時に未到達なら {@link #resolve(Map, boolean, boolean)} でフォルダ参照を省略する。
     */
    public static boolean isTaskInputSourceDirReachable(Map<String, String> ui) {
        return isDirectoryListingReachable(AppPaths.resolveTaskInputSourceDir(ui != null ? ui : Map.of()));
    }

    /** {@link #isTaskInputSourceDirReachable(Map)} と同様、実績明細ソースフォルダ用。 */
    public static boolean isActualDetailSourceDirReachable(Map<String, String> ui) {
        return isDirectoryListingReachable(AppPaths.resolveActualDetailSourceDir(ui != null ? ui : Map.of()));
    }

    private static boolean isDirectoryListingReachable(Path dir) {
        if (dir == null) {
            return false;
        }
        try {
            if (!Files.isDirectory(dir) || !Files.isReadable(dir)) {
                return false;
            }
            try (Stream<Path> s = Files.list(dir)) {
                s.findAny();
            }
            return true;
        } catch (IOException | SecurityException e) {
            return false;
        }
    }

    /** {@link Result} を merged env に適用。解決できないときは単一ファイル指定キーを外し Python 側のフォールバックに任せる。 */
    public static void applyToEnv(Map<String, String> m, Result r) {
        if (m == null || r == null) {
            return;
        }
        if (r.taskInputPath().isPresent()) {
            m.put(AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH, r.taskInputPath().get().toString());
        } else {
            m.remove(AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH);
        }
        if (r.actualDetailPath().isPresent()) {
            m.put(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK, r.actualDetailPath().get().toString());
        } else {
            m.remove(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK);
        }
    }

    static Path cacheRoot(Map<String, String> ui) {
        return AppPaths.resolveRepoRoot(ui != null ? ui : Map.of())
                .resolve(".pm-ai-cache")
                .resolve("network-source")
                .toAbsolutePath()
                .normalize();
    }

    private static Optional<Path> resolveTaskInput(
            Map<String, String> ui, List<String> logs, boolean skipSourceDirListing) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String explicit = trim(u.get(AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH));
        if (!explicit.isEmpty()) {
            Path p = Path.of(explicit).toAbsolutePath().normalize();
            if (isReadableFile(p)) {
                Optional<Path> cached = refreshCacheFromLive(p, cacheFileStemTaskInput(), u, logs);
                if (cached.isPresent()) {
                    logs.add(
                            "[network-source] 加工計画DATA相当: 参照 OK → "
                                    + p
                                    + " （キャッシュ更新）");
                }
                return Optional.of(p);
            }
            logs.add(
                    "[network-source] PM_AI_PROCESSING_PLAN_PATH が参照できません: "
                            + p
                            + " → フォルダ解決／キャッシュへフォールバックします");
        }
        if (skipSourceDirListing) {
            logs.add(
                    "[network-source] PM_AI_TASK_INPUT_SOURCE_DIR は起動時チェックで未到達のため一覧せずキャッシュを試行: "
                            + AppPaths.resolveTaskInputSourceDir(u));
            return loadTaskInputFromCache(u, logs);
        }
        Path dir = AppPaths.resolveTaskInputSourceDir(u);
        Optional<Path> live = pickNewestTaskInputInDir(dir);
        if (live.isPresent() && isReadableFile(live.get())) {
            refreshCacheFromLive(live.get(), cacheFileStemTaskInput(), u, logs);
            logs.add("[network-source] PM_AI_TASK_INPUT_SOURCE_DIR 最新: " + live.get());
            return live;
        }
        logs.add(
                "[network-source] PM_AI_TASK_INPUT_SOURCE_DIR を参照できないか空です: "
                        + dir);
        return loadTaskInputFromCache(u, logs);
    }

    private static Optional<Path> resolveActualDetail(
            Map<String, String> ui, List<String> logs, boolean skipSourceDirListing) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String wb = trim(u.get(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK));
        if (!wb.isEmpty()) {
            Path p = Path.of(wb).toAbsolutePath().normalize();
            if (isReadableFile(p)) {
                refreshCacheFromLive(p, cacheFileStemActualDetail(), u, logs);
                logs.add("[network-source] 実績明細: 単一ファイル参照 OK → " + p);
                return Optional.of(p);
            }
            logs.add(
                    "[network-source] PM_AI_ACTUAL_DETAIL_WORKBOOK が参照できません: "
                            + p
                            + " → フォルダ／キャッシュへフォールバックします");
        }
        if (skipSourceDirListing) {
            logs.add(
                    "[network-source] PM_AI_ACTUAL_DETAIL_SOURCE_DIR は起動時チェックで未到達のため一覧せずキャッシュを試行: "
                            + AppPaths.resolveActualDetailSourceDir(u));
            return loadActualDetailFromCache(u, logs);
        }
        Path dir = AppPaths.resolveActualDetailSourceDir(u);
        Optional<Path> live = pickNewestExcelInDir(dir);
        if (live.isPresent() && isReadableFile(live.get())) {
            refreshCacheFromLive(live.get(), cacheFileStemActualDetail(), u, logs);
            logs.add("[network-source] PM_AI_ACTUAL_DETAIL_SOURCE_DIR 最新: " + live.get());
            return live;
        }
        logs.add(
                "[network-source] PM_AI_ACTUAL_DETAIL_SOURCE_DIR を参照できないか空です: "
                        + dir);
        return loadActualDetailFromCache(u, logs);
    }

    private static String cacheFileStemTaskInput() {
        return "task-input-newest";
    }

    private static String cacheFileStemActualDetail() {
        return "actual-detail-newest";
    }

    private static Optional<Path> refreshCacheFromLive(
            Path liveFile, String stem, Map<String, String> ui, List<String> logs) {
        try {
            Path root = cacheRoot(ui);
            Files.createDirectories(root);
            String name = liveFile.getFileName() != null ? liveFile.getFileName().toString() : "file";
            String ext = extensionOf(name);
            Path dest = root.resolve(stem + ext);
            Files.copy(liveFile, dest, java.nio.file.StandardCopyOption.REPLACE_EXISTING);
            writeMeta(ui, stem + ext, liveFile.toString());
            return Optional.of(dest);
        } catch (IOException ex) {
            logs.add("[network-source] キャッシュ更新に失敗（無視して続行）: " + ex.getMessage());
            return Optional.empty();
        }
    }

    private static Optional<Path> loadTaskInputFromCache(Map<String, String> ui, List<String> logs) {
        return loadFromMeta(ui, cacheFileStemTaskInput(), "[network-source] 加工計画DATA相当をキャッシュから読込: ", logs);
    }

    private static Optional<Path> loadActualDetailFromCache(Map<String, String> ui, List<String> logs) {
        return loadFromMeta(ui, cacheFileStemActualDetail(), "[network-source] 実績明細をキャッシュから読込: ", logs);
    }

    private static Optional<Path> loadFromMeta(
            Map<String, String> ui, String stem, String okPrefix, List<String> logs) {
        try {
            Path root = cacheRoot(ui);
            Path metaPath = root.resolve(META_JSON);
            if (!Files.isRegularFile(metaPath)) {
                logs.add("[network-source] キャッシュメタがありません: " + metaPath);
                return Optional.empty();
            }
            String raw = Files.readString(metaPath, java.nio.charset.StandardCharsets.UTF_8);
            com.fasterxml.jackson.databind.JsonNode rootNode =
                    new com.fasterxml.jackson.databind.ObjectMapper().readTree(raw);
            com.fasterxml.jackson.databind.JsonNode slot =
                    rootNode != null ? rootNode.get(stem) : null;
            if (slot == null || !slot.isObject()) {
                logs.add("[network-source] キャッシュメタにスロットがありません: " + stem);
                return Optional.empty();
            }
            String fileName = text(slot, "cacheFile");
            if (fileName.isEmpty()) {
                return Optional.empty();
            }
            Path cached = root.resolve(fileName).toAbsolutePath().normalize();
            if (!cached.startsWith(root)) {
                logs.add("[network-source] キャッシュパスが不正です");
                return Optional.empty();
            }
            if (!isReadableFile(cached)) {
                logs.add("[network-source] キャッシュファイルが読めません: " + cached);
                return Optional.empty();
            }
            logs.add(okPrefix + cached);
            return Optional.of(cached);
        } catch (IOException ex) {
            logs.add("[network-source] キャッシュ読込エラー: " + ex.getMessage());
            return Optional.empty();
        }
    }

    private static void writeMeta(Map<String, String> ui, String cacheFileName, String sourceHint)
            throws IOException {
        Path root = cacheRoot(ui);
        Files.createDirectories(root);
        Path metaPath = root.resolve(META_JSON);
        com.fasterxml.jackson.databind.ObjectMapper om = new com.fasterxml.jackson.databind.ObjectMapper();
        com.fasterxml.jackson.databind.node.ObjectNode rootNode = om.createObjectNode();
        if (Files.isRegularFile(metaPath)) {
            try {
                com.fasterxml.jackson.databind.JsonNode prev = om.readTree(metaPath.toFile());
                if (prev != null && prev.isObject()) {
                    rootNode = (com.fasterxml.jackson.databind.node.ObjectNode) prev;
                }
            } catch (IOException ignored) {
                rootNode = om.createObjectNode();
            }
        }
        String stem = cacheFileStemFromCacheFileName(cacheFileName);
        com.fasterxml.jackson.databind.node.ObjectNode slot = om.createObjectNode();
        slot.put("cacheFile", cacheFileName);
        slot.put("sourcePath", sourceHint != null ? sourceHint : "");
        slot.put("updatedMillis", System.currentTimeMillis());
        rootNode.set(stem, slot);
        om.writerWithDefaultPrettyPrinter().writeValue(metaPath.toFile(), rootNode);
    }

    private static String cacheFileStemFromCacheFileName(String cacheFileName) {
        String n = cacheFileName != null ? cacheFileName : "";
        int dot = n.lastIndexOf('.');
        if (dot <= 0) {
            return n;
        }
        return n.substring(0, dot);
    }

    private static String text(com.fasterxml.jackson.databind.JsonNode o, String key) {
        com.fasterxml.jackson.databind.JsonNode n = o != null ? o.get(key) : null;
        if (n == null || !n.isTextual()) {
            return "";
        }
        return n.asText("").strip();
    }

    private static Optional<Path> pickNewestTaskInputInDir(Path dir) {
        if (!isAccessibleDir(dir)) {
            return Optional.empty();
        }
        try (Stream<Path> stream = Files.list(dir)) {
            return stream
                    .filter(Files::isRegularFile)
                    .filter(NetworkSourceDirResolver::isTaskInputSuffix)
                    .filter(p -> !lockFile(p))
                    .max(Comparator.comparingLong(NetworkSourceDirResolver::mtimeScore));
        } catch (IOException e) {
            return Optional.empty();
        }
    }

    private static Optional<Path> pickNewestExcelInDir(Path dir) {
        if (!isAccessibleDir(dir)) {
            return Optional.empty();
        }
        try (Stream<Path> stream = Files.list(dir)) {
            return stream
                    .filter(Files::isRegularFile)
                    .filter(NetworkSourceDirResolver::isExcelSuffix)
                    .filter(p -> !lockFile(p))
                    .max(Comparator.comparingLong(NetworkSourceDirResolver::mtimeScore));
        } catch (IOException e) {
            return Optional.empty();
        }
    }

    private static boolean isAccessibleDir(Path dir) {
        try {
            return Files.isDirectory(dir) && Files.isReadable(dir);
        } catch (Exception e) {
            return false;
        }
    }

    private static boolean isReadableFile(Path p) {
        try {
            return Files.isRegularFile(p) && Files.isReadable(p);
        } catch (Exception e) {
            return false;
        }
    }

    private static boolean lockFile(Path p) {
        String name = p.getFileName() != null ? p.getFileName().toString() : "";
        return name.startsWith("~$");
    }

    private static boolean isTaskInputSuffix(Path p) {
        String n = p.getFileName() != null ? p.getFileName().toString().toLowerCase(Locale.ROOT) : "";
        for (String s : TASK_INPUT_SUFFIXES) {
            if (n.endsWith(s)) {
                return true;
            }
        }
        return false;
    }

    private static boolean isExcelSuffix(Path p) {
        String n = p.getFileName() != null ? p.getFileName().toString().toLowerCase(Locale.ROOT) : "";
        return n.endsWith(".xlsx") || n.endsWith(".xlsm");
    }

    private static String extensionOf(String fileName) {
        int dot = fileName.lastIndexOf('.');
        return dot >= 0 ? fileName.substring(dot).toLowerCase(Locale.ROOT) : "";
    }

    private static long mtimeScore(Path p) {
        try {
            BasicFileAttributes a = Files.readAttributes(p, BasicFileAttributes.class);
            long m = a.lastModifiedTime().toMillis();
            long ac = 0L;
            try {
                ac = a.lastAccessTime().toMillis();
            } catch (UnsupportedOperationException ignored) {
                ac = 0L;
            }
            return Math.max(m, ac);
        } catch (IOException e) {
            return Long.MIN_VALUE;
        }
    }

    private static String trim(String s) {
        return s != null ? s.strip() : "";
    }
}
