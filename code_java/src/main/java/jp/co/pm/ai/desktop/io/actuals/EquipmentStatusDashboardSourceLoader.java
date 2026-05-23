package jp.co.pm.ai.desktop.io.actuals;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.io.JsonTableIo;
import jp.co.pm.ai.desktop.io.NetworkSourceFileReloadCache;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.TaskInputSourceRawGridIo;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardBuilder.ActualsSnapshot;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardBuilder.AladdinSnapshot;
import jp.co.pm.ai.desktop.io.actuals.EquipmentStatusDashboardBuilder.DispatchSnapshot;

/** ダッシュボード用3系統データのディスク読込。 */
public final class EquipmentStatusDashboardSourceLoader {

    public record LoadedSources(
            ActualsSnapshot actuals,
            AladdinSnapshot aladdin,
            DispatchSnapshot dispatch,
            String actualSourceLabel,
            String aladdinSourceLabel,
            String dispatchSourceLabel) {}

    private EquipmentStatusDashboardSourceLoader() {}

    public static LoadedSources load(Map<String, String> ui) throws IOException {
        Map<String, String> env = ui != null ? ui : Map.of();
        ActualsSnapshot actuals = loadActuals(env);
        AladdinSnapshot aladdin = loadAladdin(env);
        DispatchSnapshot dispatch = loadDispatch(env);
        return new LoadedSources(
                actuals,
                aladdin,
                dispatch,
                actualsLabel(env),
                aladdinLabel(env),
                dispatchLabel(env));
    }

    private static ActualsSnapshot loadActuals(Map<String, String> ui) throws IOException {
        NetworkSourceDirResolver.Result r = NetworkSourceDirResolver.resolve(ui);
        Optional<Path> resolved = r.actualDetailPath();
        if (resolved.isEmpty()) {
            return new ActualsSnapshot(List.of(), List.of());
        }
        Path file = resolved.get().toAbsolutePath().normalize();
        Optional<NetworkSourceFileReloadCache.Snapshot> cached =
                NetworkSourceFileReloadCache.matchActuals(file);
        if (cached.isPresent()) {
            return EquipmentStatusDashboardBuilder.actualsFrom(cached.get().toTabularSheet());
        }
        AppPaths.ensureActualDetailRawFileWithinLimit(file, ui);
        String low = file.getFileName().toString().toLowerCase(Locale.ROOT);
        if (low.endsWith(".pq") || low.endsWith(".parquet")) {
            return new ActualsSnapshot(List.of(), List.of());
        }
        int sheetIdx = 0;
        if (isExcelPath(file)) {
            List<String> names = TaskInputSourceRawGridIo.listExcelSheetNames(file);
            if (names.isEmpty()) {
                return new ActualsSnapshot(List.of(), List.of());
            }
            sheetIdx = preferredSheetIndex(names, ui);
        }
        PlanInputTabularIo.TabularSheet raw =
                TaskInputSourceRawGridIo.readRaw(file, sheetIdx, null);
        PlanInputTabularIo.TabularSheet stepped =
                TaskInputSourceRawGridIo.applyProcessingActualsDisplaySteps(raw);
        PlanInputTabularIo.TabularSheet shaped =
                TaskInputSourceRawGridIo.applyProcessingActualsDateTimeColumns(stepped);
        return EquipmentStatusDashboardBuilder.actualsFrom(shaped);
    }

    private static AladdinSnapshot loadAladdin(Map<String, String> ui) {
        Path path = AppPaths.resolveShapedAladdinPlanJsonPath(ui);
        if (!Files.isRegularFile(path)) {
            return new AladdinSnapshot(List.of(), List.of());
        }
        try {
            JsonTableIo.ArrayTable t = JsonTableIo.loadArrayTable(path);
            return new AladdinSnapshot(t.columns(), t.rows());
        } catch (Exception ex) {
            return new AladdinSnapshot(List.of(), List.of());
        }
    }

    private static DispatchSnapshot loadDispatch(Map<String, String> ui) {
        Path path = AppPaths.resolveResultDispatchTableJsonPath(ui);
        if (!Files.isRegularFile(path)) {
            return new DispatchSnapshot(List.of(), List.of());
        }
        try {
            PlanInputTabularIo.TabularSheet sheet = JsonTableIo.loadFlatTable(path).toTabularSheet();
            return EquipmentStatusDashboardBuilder.dispatchFrom(sheet);
        } catch (Exception ex) {
            return new DispatchSnapshot(List.of(), List.of());
        }
    }

    private static String actualsLabel(Map<String, String> ui) {
        NetworkSourceDirResolver.Result r = NetworkSourceDirResolver.resolve(ui);
        return r.actualDetailPath().map(p -> p.getFileName().toString()).orElse("(未設定)");
    }

    private static String aladdinLabel(Map<String, String> ui) {
        Path p = AppPaths.resolveShapedAladdinPlanJsonPath(ui);
        return Files.isRegularFile(p) ? p.getFileName().toString() : "(なし)";
    }

    private static String dispatchLabel(Map<String, String> ui) {
        Path p = AppPaths.resolveResultDispatchTableJsonPath(ui);
        return Files.isRegularFile(p) ? p.getFileName().toString() : "(なし)";
    }

    private static boolean isExcelPath(Path file) {
        String n = file.getFileName().toString().toLowerCase(Locale.ROOT);
        return n.endsWith(".xlsx") || n.endsWith(".xlsm") || n.endsWith(".xls");
    }

    private static int preferredSheetIndex(List<String> names, Map<String, String> ui) {
        if (names == null || names.isEmpty()) {
            return 0;
        }
        String want = ui != null ? ui.get(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SHEET) : null;
        if (want != null) {
            want = want.strip();
        }
        if (want == null || want.isEmpty()) {
            return 0;
        }
        int ix = names.indexOf(want);
        return ix >= 0 ? ix : 0;
    }
}
