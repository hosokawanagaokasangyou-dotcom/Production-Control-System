package jp.co.pm.ai.desktop.dispatch;

import java.nio.file.Path;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.Stage2OutputNaming;

/**
 * 表示専用のパス。選択が自分のローカル最新のときは、従来の {@link AppPaths} 解決と同一。
 * スナップショット中に解決が壊れても、書き込み用のローカル正本は変えない。
 */
public final class DispatchResultPaths {

    private DispatchResultPaths() {}

    public static Path planJson(Map<String, String> ui, DispatchResultSelection selection) {
        if (useLocal(selection)) {
            return localPlan(ui);
        }
        return newestInSnapshot(ui, selection, Stage2OutputNaming::acceptsPrimaryPlanJson);
    }

    public static Path memberJson(Map<String, String> ui, DispatchResultSelection selection) {
        if (useLocal(selection)) {
            return localMember(ui);
        }
        return newestInSnapshot(ui, selection, Stage2OutputNaming::acceptsPrimaryMemberJson);
    }

    public static Path dispatchJson(Map<String, String> ui, DispatchResultSelection selection) {
        if (useLocal(selection)) {
            return localDispatch(ui);
        }
        Path dir = snapshotDir(ui, selection);
        if (dir == null) {
            return localDispatch(ui);
        }
        return dir.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME);
    }

    public static Path shapedAladdin(Map<String, String> ui, DispatchResultSelection selection) {
        if (useLocal(selection)) {
            return AppPaths.resolveShapedAladdinPlanJsonPath(ui);
        }
        Path dir = snapshotDir(ui, selection);
        if (dir == null) {
            return AppPaths.resolveShapedAladdinPlanJsonPath(ui);
        }
        return dir.resolve(AppPaths.SHAPED_ALADDIN_PLAN_JSON_BASENAME);
    }

    public static Path shapedActuals(Map<String, String> ui, DispatchResultSelection selection) {
        if (useLocal(selection)) {
            return AppPaths.resolveShapedProcessingActualsJsonPath(ui);
        }
        Path dir = snapshotDir(ui, selection);
        if (dir == null) {
            return AppPaths.resolveShapedProcessingActualsJsonPath(ui);
        }
        return dir.resolve(AppPaths.SHAPED_PROCESSING_ACTUALS_JSON_BASENAME);
    }

    private static boolean useLocal(DispatchResultSelection selection) {
        return selection == null || selection.isLocalLatest();
    }

    private static Path snapshotDir(Map<String, String> ui, DispatchResultSelection selection) {
        try {
            return DispatchSnapshotStore.generationDir(
                    AppPaths.resolveDispatchSnapshotRoot(ui),
                    selection.operatorDir(),
                    selection.generationDir());
        } catch (RuntimeException ex) {
            return null;
        }
    }

    private static Path newestInSnapshot(
            Map<String, String> ui,
            DispatchResultSelection selection,
            java.util.function.Predicate<Path> accept) {
        Path dir = snapshotDir(ui, selection);
        if (dir == null) {
            return null;
        }
        try {
            if (!java.nio.file.Files.isDirectory(dir)) {
                return null;
            }
            Path best = null;
            try (var stream = java.nio.file.Files.list(dir)) {
                for (Path p : stream.toList()) {
                    if (!java.nio.file.Files.isRegularFile(p) || !accept.test(p)) {
                        continue;
                    }
                    if (best == null || p.getFileName().toString().compareTo(best.getFileName().toString()) > 0) {
                        best = p;
                    }
                }
            }
            return best;
        } catch (Exception ex) {
            return null;
        }
    }

    private static Path localPlan(Map<String, String> ui) {
        try {
            return Stage2OutputNaming.newestPrimaryPlanJson(AppPaths.defaultPlanningOutputDir(ui));
        } catch (Exception ex) {
            return null;
        }
    }

    private static Path localMember(Map<String, String> ui) {
        try {
            return Stage2OutputNaming.newestPrimaryMemberJson(AppPaths.defaultPlanningOutputDir(ui));
        } catch (Exception ex) {
            return null;
        }
    }

    private static Path localDispatch(Map<String, String> ui) {
        try {
            return AppPaths.resolveResultDispatchTableJsonPath(ui);
        } catch (RuntimeException ex) {
            return null;
        }
    }
}
