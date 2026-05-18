package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * 段階1実行前に planning_core の AI 備考キャッシュ（{@value #AI_REMARKS_CACHE_FILENAME}）を削除する。
 * Excel マクロの {@code AI解析_Remarksキャッシュファイルを削除} と同趣旨。
 */
public final class Stage1AiCacheClearer {

    public static final String AI_REMARKS_CACHE_FILENAME = "ai_remarks_cache.json";

    private Stage1AiCacheClearer() {}

    /** 候補パスごとに削除を試み、実行・ログ向けの行を返す。 */
    public static List<String> clearBeforeStage1Run(Map<String, String> ui) {
        List<String> logs = new ArrayList<>();
        for (Path path : candidateAiRemarksCachePaths(ui)) {
            if (!Files.isRegularFile(path)) {
                logs.add("[stage1-cache] なし: " + path);
                continue;
            }
            try {
                Files.delete(path);
                logs.add("[stage1-cache] 削除: " + path);
            } catch (IOException ex) {
                logs.add(
                        "[stage1-cache] 削除失敗: "
                                + path
                                + " — "
                                + (ex.getMessage() != null ? ex.getMessage() : ex));
            }
        }
        return logs;
    }

    static Set<Path> candidateAiRemarksCachePaths(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        Set<Path> paths = new LinkedHashSet<>();
        Path codeJsonDir = AppPaths.stage1ExcludeRulesJsonPath(u).getParent();
        if (codeJsonDir != null) {
            paths.add(codeJsonDir.resolve(AI_REMARKS_CACHE_FILENAME).toAbsolutePath().normalize());
        }
        Path pyJsonDir = AppPaths.resolvePythonScriptDir(u).resolve("json");
        paths.add(pyJsonDir.resolve(AI_REMARKS_CACHE_FILENAME).toAbsolutePath().normalize());
        paths.add(
                AppPaths.resolveDefaultOutputDir(u)
                        .resolve(AI_REMARKS_CACHE_FILENAME)
                        .toAbsolutePath()
                        .normalize());
        return paths;
    }
}
