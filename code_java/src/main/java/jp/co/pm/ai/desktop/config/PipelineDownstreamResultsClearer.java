package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Stream;

import jp.co.pm.ai.desktop.dispatch.Stage21TrialSnapshotStore;
import jp.co.pm.ai.desktop.dispatch.rules.trace.DispatchRuleTraceLoader;
import jp.co.pm.ai.desktop.io.Stage2OutputNaming;
import jp.co.pm.ai.planning.stage2.Stage2InProgressNextDayDispatchIo;

/**
 * 段階1または段階2.0 実行開始時に、段階2〜段階2.1 の成果物（配台表・段階2.1・段階2 計画/人員ブック等）をディスクから削除する。
 * 段階2.0 では「段階2実行前にキャッシュをクリアしない」が OFF のときのみ実行する。
 *
 * <p>段階1のタスク入力ブック本体（入力1表）は維持する。
 */
public final class PipelineDownstreamResultsClearer {

    private static final String DISPATCH_XLSX_BASENAME = "結果_配台表.xlsx";
    private static final String DISPATCH_SHORTFALL_BASENAME = "dispatch_trial_shortages.json";
    private static final String OVERTIME_OVERRIDES_BASENAME = "overtime_simulation_overrides.json";

    public record ClearResult(int deletedCount, int missingCount, int failedCount, List<String> detailLines) {
        public boolean anyDeleted() {
            return deletedCount > 0;
        }

        public boolean anyFailed() {
            return failedCount > 0;
        }
    }

    private PipelineDownstreamResultsClearer() {}

    /** 段階2〜段階2.1 のディスク成果物を削除する。 */
    public static ClearResult clearStage2Downstream(Map<String, String> ui) {
        return clearStage2Downstream(ui, false);
    }

    /**
     * 段階2〜段階2.1 のディスク成果物を削除する。
     *
     * @param preserveTodayDispatchSourceBundle true のとき当日配台ソース束
     *     ({@link jp.co.pm.ai.planning.stage2.source.Stage1SourceBundleIo}) は削除しない（段階2.0
     *     実行直前クリア向け）
     */
    public static ClearResult clearStage2Downstream(
            Map<String, String> ui, boolean preserveTodayDispatchSourceBundle) {
        Map<String, String> u = ui != null ? ui : Map.of();
        List<String> logs = new ArrayList<>();
        int deleted = 0;
        int missing = 0;
        int failed = 0;

        Set<Path> targets = new LinkedHashSet<>();
        collectDispatchTableArtifacts(u, targets);
        collectStage2PrimaryArtifacts(u, targets);
        targets.add(DispatchRuleTraceLoader.sidecarPath(u));
        targets.add(Stage2InProgressNextDayDispatchIo.defaultCachePath(u));
        if (!preserveTodayDispatchSourceBundle) {
            targets.add(
                    jp.co.pm.ai.planning.stage2.source.Stage1SourceBundleIo.defaultCachePath(u));
        }
        targets.add(
                jp.co.pm.ai.planning.stage2.Stage2AladdinTodayExcludeNextDayDispatchIo.defaultCachePath(
                        u));
        collectStage21DirectoryFiles(u, targets);

        for (Path path : targets) {
            if (path == null) {
                continue;
            }
            DeleteOutcome outcome = deleteRegularFile(path, logs);
            deleted += outcome.deleted();
            missing += outcome.missing();
            failed += outcome.failed();
        }

        pruneEmptyStage21Directories(u);

        logs.add("[stage1-downstream] 段階2〜段階2.1 の成果物クリアを完了しました。");
        return new ClearResult(deleted, missing, failed, List.copyOf(logs));
    }

    private static void collectDispatchTableArtifacts(Map<String, String> ui, Set<Path> targets) {
        Path dispatchJson = AppPaths.resolveResultDispatchTableJsonPath(ui);
        if (dispatchJson != null) {
            targets.add(dispatchJson);
            targets.add(dispatchJson.resolveSibling(DISPATCH_XLSX_BASENAME));
            targets.add(dispatchJson.resolveSibling(DISPATCH_SHORTFALL_BASENAME));
            targets.add(dispatchJson.resolveSibling(OVERTIME_OVERRIDES_BASENAME));
            Path stage21Sidecar = Stage21TrialSnapshotStore.sidecarPathFor(dispatchJson);
            if (stage21Sidecar != null) {
                targets.add(stage21Sidecar);
            }
        }
        targets.add(AppPaths.resolveShapedAladdinPlanJsonPath(ui));
        targets.add(AppPaths.resolveShapedProcessingActualsJsonPath(ui));
    }

    private static void collectStage2PrimaryArtifacts(Map<String, String> ui, Set<Path> targets) {
        Path outputDir = AppPaths.defaultPlanningOutputDir(ui);
        if (outputDir == null || !Files.isDirectory(outputDir)) {
            return;
        }
        try (Stream<Path> stream = Files.walk(outputDir)) {
            stream.filter(Files::isRegularFile)
                    .filter(PipelineDownstreamResultsClearer::isStage2PrimaryArtifact)
                    .forEach(targets::add);
        } catch (IOException ex) {
            // walk 失敗時は dispatch 系のみ削除（上で収集済み）
        }
    }

    private static boolean isStage2PrimaryArtifact(Path path) {
        String name = path.getFileName() != null ? path.getFileName().toString() : "";
        if (name.startsWith("~$")) {
            return false;
        }
        return Stage2OutputNaming.acceptsPrimaryPlanXlsx(path)
                || Stage2OutputNaming.acceptsPrimaryPlanJson(path)
                || Stage2OutputNaming.acceptsPrimaryMemberXlsx(path)
                || Stage2OutputNaming.acceptsPrimaryMemberJson(path);
    }

    private static void collectStage21DirectoryFiles(Map<String, String> ui, Set<Path> targets) {
        Path stage21Dir = AppPaths.resolveStage21OutputDir(ui);
        if (stage21Dir == null || !Files.isDirectory(stage21Dir)) {
            return;
        }
        try (Stream<Path> stream = Files.walk(stage21Dir)) {
            stream.filter(Files::isRegularFile).forEach(targets::add);
        } catch (IOException ignored) {
        }
    }

    private record DeleteOutcome(int deleted, int missing, int failed) {}

    private static DeleteOutcome deleteRegularFile(Path path, List<String> logs) {
        Path normalized = path.toAbsolutePath().normalize();
        if (!Files.isRegularFile(normalized)) {
            return new DeleteOutcome(0, 1, 0);
        }
        try {
            Files.delete(normalized);
            logs.add("[stage1-downstream] 削除: " + normalized);
            return new DeleteOutcome(1, 0, 0);
        } catch (IOException ex) {
            logs.add(
                    "[stage1-downstream] 削除失敗: "
                            + normalized
                            + " — "
                            + (ex.getMessage() != null ? ex.getMessage() : ex));
            return new DeleteOutcome(0, 0, 1);
        }
    }

    /** stage21 配下の空ディレクトリを削除（ファイル削除後）。 */
    private static void pruneEmptyStage21Directories(Map<String, String> ui) {
        Path stage21Dir = AppPaths.resolveStage21OutputDir(ui);
        if (stage21Dir == null || !Files.isDirectory(stage21Dir)) {
            return;
        }
        try (Stream<Path> walk = Files.walk(stage21Dir)) {
            List<Path> dirs =
                    walk.filter(Files::isDirectory)
                            .sorted(Comparator.reverseOrder())
                            .toList();
            for (Path dir : dirs) {
                if (dir.equals(stage21Dir)) {
                    continue;
                }
                try (Stream<Path> children = Files.list(dir)) {
                    if (children.findAny().isEmpty()) {
                        Files.deleteIfExists(dir);
                    }
                }
            }
        } catch (IOException ignored) {
        }
    }
}
