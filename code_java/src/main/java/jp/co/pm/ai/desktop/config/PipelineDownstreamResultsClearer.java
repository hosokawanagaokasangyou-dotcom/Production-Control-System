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

import jp.co.pm.ai.desktop.PlanInputStage3TabController;
import jp.co.pm.ai.desktop.dispatch.Stage21TrialSnapshotStore;
import jp.co.pm.ai.desktop.dispatch.Stage3PlanningMetaStore;
import jp.co.pm.ai.desktop.dispatch.rules.trace.DispatchRuleTraceLoader;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.Stage2OutputNaming;
import jp.co.pm.ai.planning.stage2.Stage2InProgressNextDayDispatchIo;

/**
 * 段階1実行開始時に、段階2〜段階3.2 の成果物（配台表・段階2.1・入力3表・段階2 計画/人員ブック等）をディスクから削除する。
 *
 * <p>段階1のタスク入力ブック本体（入力1表）は維持する。入力3表シートのみ行を空にする。
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

    /** 段階2〜3.2 のディスク成果物を削除し、入力3表シートを空にする。 */
    public static ClearResult clearStage2ThroughStage32(Map<String, String> ui) {
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

        ClearResult sheetResult = clearStage3InputSheetRows(u, logs);
        deleted += sheetResult.deletedCount();
        missing += sheetResult.missingCount();
        failed += sheetResult.failedCount();

        pruneEmptyStage21Directories(u);

        logs.add("[stage1-downstream] 段階2〜3.2 の成果物クリアを完了しました。");
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
            Path stage3Sidecar = Stage3PlanningMetaStore.sidecarPath(dispatchJson);
            if (stage3Sidecar != null) {
                targets.add(stage3Sidecar);
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

    private static ClearResult clearStage3InputSheetRows(Map<String, String> ui, List<String> logs) {
        Path workbook = AppPaths.defaultStage1PlanTasksPath(ui);
        if (workbook == null || !Files.isRegularFile(workbook)) {
            logs.add("[stage1-downstream] 入力3表クリア: タスク入力ブックなし（スキップ）");
            return new ClearResult(0, 1, 0, List.of());
        }
        try {
            PlanInputTabularIo.TabularSheet sheet =
                    PlanInputTabularIo.read(workbook, PlanInputStage3TabController.STAGE3_SHEET_NAME);
            PlanInputTabularIo.writeExcelSheetPreservingOthers(
                    workbook,
                    PlanInputStage3TabController.STAGE3_SHEET_NAME,
                    new PlanInputTabularIo.TabularSheet(sheet.headers(), List.of()));
            logs.add(
                    "[stage1-downstream] 入力3表シートを空にしました: "
                            + PlanInputStage3TabController.STAGE3_SHEET_NAME);
            return new ClearResult(1, 0, 0, List.of());
        } catch (Exception ex) {
            logs.add(
                    "[stage1-downstream] 入力3表クリア: シート未作成または失敗（"
                            + (ex.getMessage() != null ? ex.getMessage() : ex)
                            + "）");
            return new ClearResult(0, 1, 0, List.of());
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
