package jp.co.pm.ai.planning.stage2.source;

import java.nio.file.Path;
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
        if (bundle == null) {
            return Result.blocked(
                    "段階1のソース束が未保存です。当日配台 ON のときは段階1を先に正常終了してください。");
        }
        Optional<Path> currentPlan = resolveCurrentProcessingPlanPath(ui);
        Optional<Path> currentDaily = KonanDailyReportLookup.resolveNewestCsvPath(ui);
        Path expectedPlan = bundle.processingPlanPathNormalized();
        Path expectedDaily = bundle.dailyReportCsvPathNormalized();

        if (currentPlan.isEmpty()) {
            return Result.blocked("加工計画ファイルを解決できません。段階1と同じ取得分を選んでください。");
        }
        if (currentDaily.isEmpty()) {
            return Result.blocked("加工日報 CSV を解決できません。段階1と同じ取得分を選んでください。");
        }
        if (!pathsEqual(currentPlan.get(), expectedPlan)) {
            return Result.blocked(
                    "加工計画が段階1固定ソースと一致しません。\n"
                            + "固定: "
                            + expectedPlan
                            + "\n現在: "
                            + currentPlan.get()
                            + "\n段階1を再実行するか、当日配台を OFF にしてください。");
        }
        if (!pathsEqual(currentDaily.get(), expectedDaily)) {
            return Result.blocked(
                    "加工日報が段階1固定ソースと一致しません。\n"
                            + "固定: "
                            + expectedDaily
                            + "\n現在: "
                            + currentDaily.get()
                            + "\n段階1を再実行するか、当日配台を OFF にしてください。");
        }
        return Result.ok();
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
        return a.toAbsolutePath().normalize().equals(b.toAbsolutePath().normalize());
    }
}
