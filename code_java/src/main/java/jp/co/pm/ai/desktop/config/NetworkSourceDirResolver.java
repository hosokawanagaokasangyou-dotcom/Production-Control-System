package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

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

    /**
     * 加工計画の読込対象拡張子は {@link AppPaths#resolveProcessingPlanRequiredExt(Map)}。
     * 最新判定の候補拡張子は {@link AppPaths#resolveProcessingPlanCandidateExts(Map)}。
     */

    /** Office Open XML（ZIP）系。中央ディレクトリ不整合を避けるためキャッシュ時に正規化する。 */
    private static final Set<String> OFFICE_ZIP_EXTENSIONS =
            Set.of(".xlsx", ".xlsm", ".xltx", ".xltm");

    /**
     * @param taskInputFromCache {@code true} iff ネットワークソース未到達などでキャッシュへフォールバックした
     * @param actualDetailFromCache 同上（実績明細）
     */
    public record Result(
            Optional<Path> taskInputPath,
            boolean taskInputFromCache,
            Optional<Path> actualDetailPath,
            boolean actualDetailFromCache,
            List<String> logLines) {}

    /** 解決パスと、ネットワーク未到達によるキャッシュフォールバックかどうか。 */
    private record ResolvedNetworkSource(Optional<Path> path, boolean cacheFallback) {}

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
        ResolvedNetworkSource task = resolveTaskInput(m, logs, skipTaskInputSourceDirListing);
        ResolvedNetworkSource actual = resolveActualDetail(m, logs, skipActualDetailSourceDirListing);
        boolean tCache = task.cacheFallback();
        boolean aCache = actual.cacheFallback();
        return new Result(task.path(), tCache, actual.path(), aCache, List.copyOf(logs));
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

    /** {@link AppPaths#KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR} で解決される依頼書原本フォルダの一覧可否。 */
    public static boolean isRequestFormOriginalDirReachable(Map<String, String> ui) {
        return isDirectoryListingReachable(
                AppPaths.resolveRequestFormOriginalDir(ui != null ? ui : Map.of()));
    }

    /** {@link AppPaths#KEY_PM_AI_REQUEST_FORM_TPI_PDF_DIR} で解決される TPI PDF フォルダの一覧可否。未設定時は {@code false}。 */
    public static boolean isRequestFormTpiPdfDirReachable(Map<String, String> ui) {
        return AppPaths.resolveRequestFormTpiPdfDir(ui != null ? ui : Map.of())
                .map(NetworkSourceDirResolver::isDirectoryListingReachable)
                .orElse(false);
    }

    /**
     * ディレクトリとして存在し、{@link Files#list} が成功するか（UNC 未到達等は {@code false}）。
     */
    public static boolean isDirectoryListingReachable(Path dir) {
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

    private static ResolvedNetworkSource resolveTaskInput(
            Map<String, String> ui, List<String> logs, boolean skipSourceDirListing) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String explicit = trim(u.get(AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH));
        if (!explicit.isEmpty()) {
            Path p = Path.of(explicit).toAbsolutePath().normalize();
            if (isReadableFile(p)) {
                SourceFileExtensionPolicy.Result ext =
                        SourceFileExtensionPolicy.checkProcessingPlanFile(p, u);
                if (!ext.ok()) {
                    logs.add("[network-source] " + ext.errorMessage());
                    return new ResolvedNetworkSource(Optional.empty(), false);
                }
                Optional<Path> cached = refreshCacheFromLive(p, cacheFileStemTaskInput(), u, logs);
                if (cached.isPresent()) {
                    logs.add(
                            "[network-source] 加工計画DATA相当: 参照 OK → "
                                    + p
                                    + " （ローカルキャッシュ読込: "
                                    + cached.get()
                                    + "）");
                    return new ResolvedNetworkSource(cached, false);
                }
                logs.add("[network-source] 加工計画DATA相当: 参照 OK → " + p);
                return new ResolvedNetworkSource(Optional.of(p), false);
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
            return cacheFallbackResult(loadTaskInputFromCache(u, logs));
        }
        Path dir = AppPaths.resolveTaskInputSourceDir(u);
        SourceFileExtensionPolicy.Result planExt =
                SourceFileExtensionPolicy.checkProcessingPlanDirectory(dir, u);
        if (!planExt.ok()) {
            logs.add("[network-source] " + planExt.errorMessage());
            if (planExt.newestCandidatePath().isPresent()) {
                return new ResolvedNetworkSource(Optional.empty(), false);
            }
            return cacheFallbackResult(loadTaskInputFromCache(u, logs));
        }
        Optional<Path> live = planExt.loadablePath();
        if (live.isPresent() && isReadableFile(live.get())) {
            Optional<Path> cached =
                    refreshCacheFromLive(live.get(), cacheFileStemTaskInput(), u, logs);
            if (cached.isPresent()) {
                logs.add(
                        "[network-source] PM_AI_TASK_INPUT_SOURCE_DIR 最新（ローカルキャッシュ）: "
                                + cached.get());
                return new ResolvedNetworkSource(cached, false);
            }
            logs.add("[network-source] PM_AI_TASK_INPUT_SOURCE_DIR 最新: " + live.get());
            return new ResolvedNetworkSource(live, false);
        }
        logs.add(
                "[network-source] PM_AI_TASK_INPUT_SOURCE_DIR を参照できないか空です: "
                        + dir);
        return cacheFallbackResult(loadTaskInputFromCache(u, logs));
    }

    private static ResolvedNetworkSource resolveActualDetail(
            Map<String, String> ui, List<String> logs, boolean skipSourceDirListing) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String wb = trim(u.get(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_WORKBOOK));
        if (!wb.isEmpty()) {
            Path p = Path.of(wb).toAbsolutePath().normalize();
            if (isReadableFile(p)) {
                Optional<Path> cached = refreshCacheFromLive(p, cacheFileStemActualDetail(), u, logs);
                if (cached.isPresent()) {
                    logs.add(
                            "[network-source] 実績明細: 単一ファイル参照 OK → "
                                    + p
                                    + " （ローカルキャッシュ読込: "
                                    + cached.get()
                                    + "）");
                    return new ResolvedNetworkSource(cached, false);
                }
                logs.add("[network-source] 実績明細: 単一ファイル参照 OK → " + p);
                return new ResolvedNetworkSource(Optional.of(p), false);
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
            return cacheFallbackResult(loadActualDetailFromCache(u, logs));
        }
        Path dir = AppPaths.resolveActualDetailSourceDir(u);
        Optional<Path> live = pickNewestExcelInDir(dir);
        if (live.isPresent() && isReadableFile(live.get())) {
            Optional<Path> cached =
                    refreshCacheFromLive(live.get(), cacheFileStemActualDetail(), u, logs);
            if (cached.isPresent()) {
                logs.add(
                        "[network-source] PM_AI_ACTUAL_DETAIL_SOURCE_DIR 最新（ローカルキャッシュ）: "
                                + cached.get());
                return new ResolvedNetworkSource(cached, false);
            }
            logs.add("[network-source] PM_AI_ACTUAL_DETAIL_SOURCE_DIR 最新: " + live.get());
            return new ResolvedNetworkSource(live, false);
        }
        logs.add(
                "[network-source] PM_AI_ACTUAL_DETAIL_SOURCE_DIR を参照できないか空です: "
                        + dir);
        return cacheFallbackResult(loadActualDetailFromCache(u, logs));
    }

    private static ResolvedNetworkSource cacheFallbackResult(Optional<Path> path) {
        return new ResolvedNetworkSource(path, path.isPresent());
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
            copyLiveFileToCache(liveFile, dest, logs);
            pruneSiblingCacheFiles(root, stem, dest.getFileName().toString());
            writeMeta(ui, stem + ext, liveFile.toString());
            return Optional.of(dest);
        } catch (IOException ex) {
            logs.add("[network-source] キャッシュ更新に失敗（無視して続行）: " + ex.getMessage());
            return Optional.empty();
        }
    }

    /**
     * ネットワーク元をローカルキャッシュへ複製する。xlsx 系は ZIP を正規化してから書く。
     * 正規化に失敗したときだけ素の {@link Files#copy} にフォールバックする。
     */
    static void copyLiveFileToCache(Path liveFile, Path dest, List<String> logs) throws IOException {
        String ext = extensionOf(dest.getFileName() != null ? dest.getFileName().toString() : "");
        if (OFFICE_ZIP_EXTENSIONS.contains(ext)) {
            try {
                rewriteOfficeZipToCache(liveFile, dest);
                return;
            } catch (IOException ex) {
                if (logs != null) {
                    logs.add(
                            "[network-source] xlsx ZIP正規化に失敗したため素コピーします: "
                                    + ex.getMessage());
                }
            }
        }
        Files.copy(liveFile, dest, StandardCopyOption.REPLACE_EXISTING);
    }

    /** ローカルヘッダから読み直し、標準的な中央ディレクトリ付き ZIP として書き出す。 */
    static void rewriteOfficeZipToCache(Path liveFile, Path dest) throws IOException {
        Path tmp = dest.resolveSibling(dest.getFileName().toString() + ".normalize.tmp");
        try {
            try (InputStream in = Files.newInputStream(liveFile);
                    ZipInputStream zin = new ZipInputStream(in);
                    OutputStream out = Files.newOutputStream(tmp);
                    ZipOutputStream zout = new ZipOutputStream(out)) {
                byte[] buf = new byte[8192];
                ZipEntry entry;
                int count = 0;
                while ((entry = zin.getNextEntry()) != null) {
                    if (entry.isDirectory()) {
                        zin.closeEntry();
                        continue;
                    }
                    String entryName = entry.getName();
                    if (entryName == null || entryName.isBlank()) {
                        zin.closeEntry();
                        continue;
                    }
                    ZipEntry outEntry = new ZipEntry(entryName);
                    zout.putNextEntry(outEntry);
                    int n;
                    while ((n = zin.read(buf)) >= 0) {
                        zout.write(buf, 0, n);
                    }
                    zout.closeEntry();
                    zin.closeEntry();
                    count++;
                }
                if (count == 0) {
                    throw new IOException("ZIPエントリがありません: " + liveFile);
                }
            }
            Files.move(tmp, dest, StandardCopyOption.REPLACE_EXISTING);
        } finally {
            try {
                Files.deleteIfExists(tmp);
            } catch (IOException ignored) {
                // best effort
            }
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

    /** 同一 stem の旧拡張子キャッシュ（例: csv → xlsx 更新後の csv）を削除する。 */
    static void pruneSiblingCacheFiles(Path root, String stem, String keepFileName) {
        if (root == null || stem == null || stem.isBlank() || keepFileName == null || keepFileName.isBlank()) {
            return;
        }
        if (!Files.isDirectory(root)) {
            return;
        }
        try (Stream<Path> stream = Files.list(root)) {
            for (Path path : stream.toList()) {
                if (!Files.isRegularFile(path)) {
                    continue;
                }
                String name = path.getFileName() != null ? path.getFileName().toString() : "";
                if (name.equals(keepFileName)) {
                    continue;
                }
                if (stem.equals(cacheFileStemFromCacheFileName(name))) {
                    Files.deleteIfExists(path);
                }
            }
        } catch (IOException ignored) {
        }
    }

    private static String text(com.fasterxml.jackson.databind.JsonNode o, String key) {
        com.fasterxml.jackson.databind.JsonNode n = o != null ? o.get(key) : null;
        if (n == null || !n.isTextual()) {
            return "";
        }
        return n.asText("").strip();
    }

    /**
     * {@link AppPaths#resolveTaskInputSourceDir(Map)} 直下から加工計画の最新を返す。
     * 最新候補の拡張子が必須拡張子以外のときは empty（{@link SourceFileExtensionPolicy}）。
     *
     * <p>{@code PM_AI_PROCESSING_PLAN_PATH} で単一ファイルが指定されている場合の優先は行わない（フォルダ内の最新のみ）。
     * 環境変数による必須拡張子は未適用（既定）。UI マップがあるときは
     * {@link #newestTaskInputFileInDirectory(Path, Map)} を使うこと。
     */
    public static Optional<Path> newestTaskInputFileInDirectory(Path taskInputSourceDir) {
        return newestTaskInputFileInDirectory(taskInputSourceDir, Map.of());
    }

    /** {@link #newestTaskInputFileInDirectory(Path)} と同じだが、必須／候補拡張子と明示パスを ui から解決する。 */
    public static Optional<Path> newestTaskInputFileInDirectory(
            Path taskInputSourceDir, Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String explicit = trim(u.get(AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH));
        if (!explicit.isEmpty()) {
            Path p = Path.of(explicit);
            if (isReadableFile(p)) {
                return SourceFileExtensionPolicy.checkProcessingPlanFile(p, u).loadablePath();
            }
        }
        return SourceFileExtensionPolicy.checkProcessingPlanDirectory(taskInputSourceDir, u)
                .loadablePath();
    }

    /** {@link AppPaths#resolveActualDetailSourceDir(Map)} 直下の最新 Excel（xlsx/xlsm）。 */
    public static Optional<Path> newestExcelFileInDirectory(Path actualDetailSourceDir) {
        return pickNewestExcelInDir(actualDetailSourceDir);
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
            return a.lastModifiedTime().toMillis();
        } catch (IOException e) {
            return Long.MIN_VALUE;
        }
    }

    private static String trim(String s) {
        return s != null ? s.strip() : "";
    }
}
