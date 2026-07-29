package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;

/**
 * PM_AI_TASK_INPUT_SOURCE_DIR から最新のアラジン加工計画をディスクから再読込し、
 * {@link AppPaths#resolveShapedAladdinPlanJsonPath} へ shaped JSON を保存する。
 *
 * <p>アラジン加工計画タブの {@code reloadFromSourceDir} と同等のソース読込だが、
 * {@link NetworkSourceFileReloadCache} の同一ファイル名キャッシュは使わず常に POI 等で再読込する。
 */
public final class AladdinProcessingPlanSourceReloader {

    /** 再読込結果（ソースファイルと shaped JSON 保存先）。 */
    public record ReloadResult(
            Path sourceFile, Path shapedJsonPath, int rowCount, int columnCount) {}

    private AladdinProcessingPlanSourceReloader() {}

    /**
     * タスク入力ソースの最新ファイルからアラジン加工計画を再読込し shaped JSON を更新する。
     *
     * @throws IOException ソース未設定・ファイル未存在・読込失敗
     */
    public static ReloadResult reloadNewestFromDiskAndSaveShapedJson(Map<String, String> ui)
            throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path dir = AppPaths.resolveTaskInputSourceDir(u);
        if (dir == null || !Files.isDirectory(dir)) {
            throw new IOException("タスク入力ソースフォルダがありません: " + dir);
        }
        Optional<Path> newest = NetworkSourceDirResolver.newestTaskInputFileInDirectory(dir);
        if (newest.isEmpty()) {
            throw new IOException("読込対象ファイルがありません: " + dir);
        }
        Path file = newest.get().toAbsolutePath().normalize();
        String low = file.getFileName().toString().toLowerCase(Locale.ROOT);
        if (low.endsWith(".pq") || low.endsWith(".parquet")) {
            throw new IOException("Parquet は未対応です: " + file);
        }

        PlanInputTabularIo.TabularSheet tab = readNewestAladdinTabularFromDisk(file);
        Path shapedJson = AppPaths.resolveShapedAladdinPlanJsonPath(u);
        Files.createDirectories(shapedJson.getParent());
        JsonTableIo.saveArrayTable(shapedJson, tab.headers(), tab.rows());
        return new ReloadResult(file, shapedJson, tab.rows().size(), tab.headers().size());
    }

    /**
     * タスク入力ソースの最新ファイルからアラジン加工計画を読み込む（shaped JSON は更新しない）。
     *
     * @throws IOException 読込失敗
     */
    public static PlanInputTabularIo.TabularSheet readNewestAladdinTabularFromDisk(Path file)
            throws IOException {
        if (file == null) {
            throw new IOException("読込対象ファイルがありません");
        }
        Path normalized = file.toAbsolutePath().normalize();
        String low = normalized.getFileName().toString().toLowerCase(Locale.ROOT);
        if (low.endsWith(".pq") || low.endsWith(".parquet")) {
            throw new IOException("Parquet は未対応です: " + normalized);
        }
        return readAladdinTabularFromDisk(normalized);
    }

    private static PlanInputTabularIo.TabularSheet readAladdinTabularFromDisk(Path file)
            throws IOException {
        if (isExcelPath(file)) {
            List<String> sheetNames = TaskInputSourceRawGridIo.listExcelSheetNames(file);
            if (sheetNames.isEmpty()) {
                throw new IOException("Excel シートがありません: " + file);
            }
            int sheetIndex = 0;
            PlanInputTabularIo.TabularSheet tab =
                    TaskInputSourceRawGridIo.applyAladdinProcessingPlanDisplaySteps(
                            TaskInputSourceRawGridIo.readRaw(file, sheetIndex));
            NetworkSourceFileReloadCache.storeAladdin(
                    file, true, new ArrayList<>(sheetNames), sheetIndex, tab);
            return tab;
        }
        PlanInputTabularIo.TabularSheet tab =
                TaskInputSourceRawGridIo.applyAladdinProcessingPlanDisplaySteps(
                        TaskInputSourceRawGridIo.readRaw(file, 0));
        NetworkSourceFileReloadCache.storeAladdin(file, false, List.of(), 0, tab);
        return tab;
    }

    private static boolean isExcelPath(Path p) {
        if (p == null || p.getFileName() == null) {
            return false;
        }
        String low = p.getFileName().toString().toLowerCase(Locale.ROOT);
        return low.endsWith(".xlsx")
                || low.endsWith(".xlsm")
                || low.endsWith(".xltx")
                || low.endsWith(".xltm");
    }
}
