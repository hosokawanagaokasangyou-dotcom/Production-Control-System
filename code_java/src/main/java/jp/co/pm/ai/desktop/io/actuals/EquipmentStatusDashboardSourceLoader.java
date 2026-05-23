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

    /** 実績・アラジン・配台のソースファイル指紋（変更検知用）。 */
    public record SourceFingerprint(String actualKey, String aladdinKey, String dispatchKey) {

        public static SourceFingerprint empty() {
            return new SourceFingerprint("", "", "");
        }
    }

    /** {@link #load} の結果。{@code sourcesUnchanged} のとき再読込は行っていない。 */
    public record ReloadDecision(
            boolean sourcesUnchanged, LoadedSources sources, SourceFingerprint fingerprint) {

        public static ReloadDecision skip() {
            return new ReloadDecision(true, null, null);
        }

        public static ReloadDecision loaded(LoadedSources sources, SourceFingerprint fingerprint) {
            return new ReloadDecision(false, sources, fingerprint);
        }
    }

    public record LoadedSources(
            ActualsSnapshot actuals,
            AladdinSnapshot aladdin,
            DispatchSnapshot dispatch,
            String actualSourceLabel,
            String aladdinSourceLabel,
            String dispatchSourceLabel) {}

    private EquipmentStatusDashboardSourceLoader() {}

    public static LoadedSources load(Map<String, String> ui) throws IOException {
        ReloadDecision d = loadIfChanged(ui, null, false);
        return d.sources();
    }

    /**
     * 指紋が前回と同一ならディスク読込を省略する。
     *
     * @param previous 前回成功時の指紋（初回は {@code null}）
     * @param haveCachedData {@code true} のときのみ省略判定（メモリ上に表示用データがある）
     */
    public static ReloadDecision loadIfChanged(
            Map<String, String> ui, SourceFingerprint previous, boolean haveCachedData)
            throws IOException {
        SourceFingerprint fp = fingerprint(ui);
        if (haveCachedData && previous != null && previous.equals(fp)) {
            return ReloadDecision.skip();
        }
        return ReloadDecision.loaded(loadSources(ui), fp);
    }

    /** 3系統のソースファイル指紋（解決パス + {@code lastModified} + サイズ + シート名）。 */
    public static SourceFingerprint fingerprint(Map<String, String> ui) {
        Map<String, String> env = ui != null ? ui : Map.of();
        String sheet = env.getOrDefault(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SHEET, "").strip();
        NetworkSourceDirResolver.Result r = NetworkSourceDirResolver.resolve(env);
        String actualKey =
                fileKey(r.actualDetailPath().orElse(null), sheet.isEmpty() ? "0" : sheet);
        return new SourceFingerprint(
                actualKey,
                aladdinFingerprintKey(env),
                fileKey(AppPaths.resolveResultDispatchTableJsonPath(env), ""));
    }

    private static String aladdinFingerprintKey(Map<String, String> env) {
        NetworkSourceDirResolver.Result r = NetworkSourceDirResolver.resolve(env);
        Optional<Path> taskInput = r.taskInputPath();
        if (taskInput.isPresent()) {
            return fileKey(taskInput.get(), "task-input|0");
        }
        Path shaped = AppPaths.resolveShapedAladdinPlanJsonPath(env);
        if (Files.isRegularFile(shaped) && shapedHasRows(shaped)) {
            return fileKey(shaped, "shaped");
        }
        return fileKey(shaped, "missing");
    }

    private static boolean shapedHasRows(Path shaped) {
        try {
            JsonTableIo.ArrayTable t = JsonTableIo.loadArrayTable(shaped);
            return !t.columns().isEmpty() || !t.rows().isEmpty();
        } catch (Exception ex) {
            return false;
        }
    }

    private static LoadedSources loadSources(Map<String, String> env) throws IOException {
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

    private static String fileKey(Path path, String suffix) {
        if (path == null) {
            return "|missing|" + nz(suffix);
        }
        Path abs = path.toAbsolutePath().normalize();
        if (!Files.isRegularFile(abs)) {
            return abs + "|missing|" + nz(suffix);
        }
        try {
            long mod = Files.getLastModifiedTime(abs).toMillis();
            long size = Files.size(abs);
            return abs + "|" + mod + "|" + size + "|" + nz(suffix);
        } catch (IOException ex) {
            return abs + "|error|" + nz(suffix);
        }
    }

    private static String nz(String s) {
        return s != null ? s : "";
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

    private static AladdinSnapshot loadAladdin(Map<String, String> ui) throws IOException {
        NetworkSourceDirResolver.Result r = NetworkSourceDirResolver.resolve(ui);
        Optional<Path> resolved = r.taskInputPath();
        if (resolved.isPresent()) {
            return loadAladdinFromFile(resolved.get());
        }
        Path shapedPath = AppPaths.resolveShapedAladdinPlanJsonPath(ui);
        if (Files.isRegularFile(shapedPath)) {
            try {
                JsonTableIo.ArrayTable t = JsonTableIo.loadArrayTable(shapedPath);
                if (!t.columns().isEmpty() || !t.rows().isEmpty()) {
                    return new AladdinSnapshot(t.columns(), t.rows());
                }
            } catch (Exception ex) {
                // ignore
            }
        }
        return new AladdinSnapshot(List.of(), List.of());
    }

    private static AladdinSnapshot loadAladdinFromFile(Path file) throws IOException {
        Path normalized = file.toAbsolutePath().normalize();
        Optional<NetworkSourceFileReloadCache.Snapshot> cached =
                NetworkSourceFileReloadCache.matchAladdin(normalized);
        if (cached.isPresent()) {
            return EquipmentStatusDashboardBuilder.aladdinFrom(cached.get().toTabularSheet());
        }
        String low = normalized.getFileName().toString().toLowerCase(Locale.ROOT);
        if (low.endsWith(".pq") || low.endsWith(".parquet")) {
            return new AladdinSnapshot(List.of(), List.of());
        }
        PlanInputTabularIo.TabularSheet raw = TaskInputSourceRawGridIo.readRaw(normalized, 0, null);
        PlanInputTabularIo.TabularSheet shaped =
                TaskInputSourceRawGridIo.applyAladdinProcessingPlanDisplaySteps(raw);
        return EquipmentStatusDashboardBuilder.aladdinFrom(shaped);
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
        return NetworkSourceDirResolver.resolve(ui)
                .taskInputPath()
                .map(p -> p.getFileName().toString())
                .orElseGet(
                        () -> {
                            Path shaped = AppPaths.resolveShapedAladdinPlanJsonPath(ui);
                            if (Files.isRegularFile(shaped) && shapedHasRows(shaped)) {
                                return shaped.getFileName().toString() + " (キャッシュJSON)";
                            }
                            return "(なし)";
                        });
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
