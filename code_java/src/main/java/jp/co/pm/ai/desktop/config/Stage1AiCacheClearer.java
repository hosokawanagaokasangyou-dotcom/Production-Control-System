package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import jp.co.pm.ai.desktop.io.NetworkSourceFileReloadCache;

/**
 * 段階1「キャッシュをクリアして実行」で削除するワークスペースキャッシュ。
 *
 * <ul>
 *   <li>AI 備考: {@value #AI_REMARKS_CACHE_FILENAME}（Excel マクロの AI 解析キャッシュ削除と同趣旨）
 *   <li>配台・納期: {@link AppPaths#RESULT_DISPATCH_TABLE_JSON_BASENAME}、{@link
 *       AppPaths#SHAPED_ALADDIN_PLAN_JSON_BASENAME}、{@link AppPaths#SHAPED_PROCESSING_ACTUALS_JSON_BASENAME}
 *   <li>メモリ: {@link NetworkSourceFileReloadCache}（アラジン計画／実績明細の同一ファイル名再読込省略）
 * </ul>
 */
public final class Stage1AiCacheClearer {

    public static final String AI_REMARKS_CACHE_FILENAME = "ai_remarks_cache.json";

    /** {@link #clearBeforeStage1Run(Map)} の結果（実行・ログ向け）。 */
    public record ClearResult(int deletedCount, int missingCount, int failedCount, List<String> detailLines) {
        public boolean anyDeleted() {
            return deletedCount > 0;
        }

        public boolean anyFailed() {
            return failedCount > 0;
        }
    }

    private Stage1AiCacheClearer() {}

    /** ディスク上のいずれかの段階1関連キャッシュが実在するか。 */
    public static boolean hasAnyExistingDiskCache(Map<String, String> ui) {
        for (Path path : allClearableDiskCachePaths(ui)) {
            if (Files.isRegularFile(path)) {
                return true;
            }
        }
        return false;
    }

    public static boolean hasExistingAiRemarksCache(Map<String, String> ui) {
        for (Map.Entry<String, Path> e : roleToActivePathMap(ui).entrySet()) {
            if (!e.getKey().startsWith("ai_remarks_")) {
                continue;
            }
            if (Files.isRegularFile(e.getValue())) {
                return true;
            }
        }
        return false;
    }

    public static boolean hasExistingWorkspaceShapedCache(Map<String, String> ui) {
        for (Map.Entry<String, Path> e : roleToActivePathMap(ui).entrySet()) {
            if (e.getKey().startsWith("ai_remarks_")) {
                continue;
            }
            if (Files.isRegularFile(e.getValue())) {
                return true;
            }
        }
        return false;
    }

    /** 「キャッシュを使用します」用: 存在するキャッシュ種別の短い説明（空ならディスクキャッシュ無し）。 */
    public static List<String> existingCacheKindLabelsJa(Map<String, String> ui) {
        List<String> kinds = new ArrayList<>();
        if (hasExistingAiRemarksCache(ui)) {
            kinds.add("AI 備考解析");
        }
        if (hasExistingWorkspaceShapedCache(ui)) {
            kinds.add("配台表・納期ビュー・アラジン計画");
        }
        return kinds;
    }

    /**
     * 復元用ロール → 現在の解決パス（ファイルが無くてもパスは返す）。
     */
    public static Map<String, Path> roleToActivePathMap(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        Map<String, Path> map = new LinkedHashMap<>();
        Path codeJsonDir = AppPaths.stage1ExcludeRulesJsonPath(u).getParent();
        if (codeJsonDir != null) {
            map.put(
                    "ai_remarks_code_json",
                    codeJsonDir.resolve(AI_REMARKS_CACHE_FILENAME).toAbsolutePath().normalize());
        }
        Path pyJsonDir = AppPaths.resolvePythonScriptDir(u).resolve("json");
        map.put(
                "ai_remarks_python_json",
                pyJsonDir.resolve(AI_REMARKS_CACHE_FILENAME).toAbsolutePath().normalize());
        map.put(
                "ai_remarks_output_legacy",
                AppPaths.resolveDefaultOutputDir(u)
                        .resolve(AI_REMARKS_CACHE_FILENAME)
                        .toAbsolutePath()
                        .normalize());
        map.put("result_dispatch", AppPaths.resolveResultDispatchTableJsonPath(u));
        map.put("shaped_aladdin_plan", AppPaths.resolveShapedAladdinPlanJsonPath(u));
        map.put(
                "shaped_processing_actuals",
                AppPaths.resolveShapedProcessingActualsJsonPath(u));
        return Map.copyOf(map);
    }

    /** 退避してからディスクキャッシュを削除し、メモリキャッシュも破棄する。 */
    public static ClearResult archiveAndClearBeforeStage1Run(
            Map<String, String> ui, String archiveLabel) throws IOException {
        List<String> logs = new ArrayList<>();
        WorkspaceCacheArchiveStore.WorkspaceCacheArchiveEntry archived =
                WorkspaceCacheArchiveStore.archiveDiskCaches(
                        ui, archiveLabel, "stage1-clear");
        if (archived != null) {
            logs.add(
                    "[stage1-cache] キャッシュを退避しました（履歴 ID: "
                            + archived.id()
                            + "、"
                            + archived.fileCount()
                            + " ファイル）。");
        }
        ClearResult cleared = clearDiskCachesOnly(ui);
        logs.addAll(cleared.detailLines());
        return new ClearResult(
                cleared.deletedCount(),
                cleared.missingCount(),
                cleared.failedCount(),
                List.copyOf(logs));
    }

    /** 候補パスごとに削除を試み、メモリキャッシュも破棄する（退避なし）。 */
    public static ClearResult clearBeforeStage1Run(Map<String, String> ui) {
        return clearDiskCachesOnly(ui);
    }

    private static ClearResult clearDiskCachesOnly(Map<String, String> ui) {
        List<String> logs = new ArrayList<>();
        int deleted = 0;
        int missing = 0;
        int failed = 0;
        for (Path path : allClearableDiskCachePaths(ui)) {
            if (!Files.isRegularFile(path)) {
                missing++;
                logs.add("[stage1-cache] なし: " + path);
                continue;
            }
            try {
                Files.delete(path);
                deleted++;
                logs.add("[stage1-cache] 削除: " + path);
            } catch (IOException ex) {
                failed++;
                logs.add(
                        "[stage1-cache] 削除失敗: "
                                + path
                                + " — "
                                + (ex.getMessage() != null ? ex.getMessage() : ex));
            }
        }
        NetworkSourceFileReloadCache.clearAll();
        logs.add("[stage1-cache] メモリ上の再読込キャッシュ（アラジン計画／実績明細）を破棄しました。");
        return new ClearResult(deleted, missing, failed, List.copyOf(logs));
    }

    static Set<Path> allClearableDiskCachePaths(Map<String, String> ui) {
        return new LinkedHashSet<>(roleToActivePathMap(ui).values());
    }
}

