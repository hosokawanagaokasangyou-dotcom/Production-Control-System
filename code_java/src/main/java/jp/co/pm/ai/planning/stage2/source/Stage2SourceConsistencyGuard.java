package jp.co.pm.ai.planning.stage2.source;

import java.nio.file.Path;
import java.nio.file.Files;
import java.util.Map;
import java.util.Optional;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.reconciliation.KonanDailyReportLookup;

/** 段階2実行前: 段階1固定 bundle と現在解決されるソースパスを照合する。 */
public final class Stage2SourceConsistencyGuard {

    public record Result(boolean allowed, String message) {
        public static Result ok() {
            return new Result(true, "");
        }

        public static Result blocked(String message) {
            return new Result(false, message != null ? message : "ソース束が一致しません");
        }
    }

    private Stage2SourceConsistencyGuard() {}

    public static Result verify(Map<String, String> ui, Stage1SourceBundle bundle) {
        return prepareForStageRun(
                true, new java.util.HashMap<>(ui != null ? ui : Map.of()), bundle);
    }

    public static Result prepareForStageRun(
            boolean todayDispatch, Map<String, String> ui, Stage1SourceBundle bundle) {
        if (!todayDispatch) {
            return Result.ok();
        }
        if (bundle == null) {
            return Result.blocked(
                    "段階1のソース束が未保存です。「当日配台する」のときは、その状態で"
                            + "段階1を実行し（ソース選択ダイアログで加工計画を選択）、正常終了させてから段階2を実行してください。");
        }
        var structural = bundle.validationError();
        if (structural.isPresent()) return Result.blocked("段階1bundleが不正です: " + structural.get());
        Path expectedPlan = bundle.processingPlanPathNormalized();
        Path expectedDaily = bundle.dailyReportCsvPathNormalized();
        Path expectedExtraction = Path.of(bundle.dataExtractionWorkbookPath()).toAbsolutePath().normalize();
        Result files = requireFiles(expectedPlan, expectedDaily, expectedExtraction);
        if (!files.allowed()) return files;
        Result explicitPlan = verifyExplicit(ui, AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH, expectedPlan, "加工計画");
        if (!explicitPlan.allowed()) return explicitPlan;
        Result explicitDaily = verifyExplicit(ui, KonanDailyReportLookup.KEY_DAILY_REPORT_CSV_PATH, expectedDaily, "加工日報");
        if (!explicitDaily.allowed()) return explicitDaily;
        overlayBundlePaths(ui, bundle);
        return Result.ok();
    }

    private static Result requireFiles(Path plan, Path daily, Path extraction) {
        if (!Files.isRegularFile(plan)) return Result.blocked("固定 processingPlanPath を利用できません");
        if (!Files.isRegularFile(daily)) return Result.blocked("固定 dailyReportCsvPath を利用できません");
        if (!Files.isRegularFile(extraction)) return Result.blocked("固定 dataExtractionWorkbookPath を利用できません");
        return Result.ok();
    }

    private static Result verifyExplicit(Map<String, String> ui, String key, Path expected, String label) {
        String raw = ui != null ? ui.getOrDefault(key, "").strip() : "";
        if (raw.isEmpty()) return Result.ok();
        final Path current;
        try { current = Path.of(raw); } catch (RuntimeException ex) { return Result.blocked(label + "の明示パスが不正です"); }
        return pathsEqual(current, expected) ? Result.ok() : Result.blocked(label + "が段階1固定ソースと一致しません");
    }

    public static void overlayBundlePaths(Map<String, String> env, Stage1SourceBundle bundle) {
        if (env == null || bundle == null) {
            return;
        }
        env.put(AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH, bundle.processingPlanPath());
        env.put(
                AppPaths.KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK,
                bundle.dataExtractionWorkbookPath());
        env.put(KonanDailyReportLookup.KEY_DAILY_REPORT_CSV_PATH, bundle.dailyReportCsvPath());
    }

    private static Optional<Path> resolveCurrentProcessingPlanPath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String explicit = u.getOrDefault(AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH, "").strip();
        if (!explicit.isEmpty()) {
            Path p = Path.of(explicit).toAbsolutePath().normalize();
            if (java.nio.file.Files.isRegularFile(p)) {
                return Optional.of(p);
            }
        }
        boolean taskReach = NetworkSourceDirResolver.isTaskInputSourceDirReachable(u);
        boolean actReach = NetworkSourceDirResolver.isActualDetailSourceDirReachable(u);
        NetworkSourceDirResolver.Result res =
                NetworkSourceDirResolver.resolve(u, !taskReach, !actReach);
        return res.taskInputPath().map(p -> p.toAbsolutePath().normalize());
    }

    private static boolean pathsEqual(Path a, Path b) {
        if (a == null || b == null) {
            return false;
        }
        try {
            if (Files.exists(a) && Files.exists(b) && Files.isSameFile(a, b)) return true;
        } catch (Exception ignored) {
        }
        try {
            return a.toAbsolutePath().normalize().equals(b.toAbsolutePath().normalize());
        } catch (RuntimeException ex) {
            return false;
        }
    }
}
