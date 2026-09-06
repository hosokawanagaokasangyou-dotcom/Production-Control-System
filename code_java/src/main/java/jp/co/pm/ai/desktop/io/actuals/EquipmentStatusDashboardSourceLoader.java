package jp.co.pm.ai.desktop.io.actuals;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.Set;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.io.JsonTableIo;
import jp.co.pm.ai.desktop.io.NetworkSourceFileReloadCache;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.TaskInputSourceRawGridIo;

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

    /** 読込時間・ソースファイル合計サイズ・行数。 */
    public record LoadStats(
            long totalSourceBytes,
            long loadDurationMs,
            int actualRowCount,
            int aladdinRowCount,
            int dispatchRowCount) {

        public static LoadStats empty() {
            return new LoadStats(0L, 0L, 0, 0, 0);
        }
    }

    public record LoadedSources(
            ActualsSnapshot actuals,
            AladdinSnapshot aladdin,
            DispatchSnapshot dispatch,
            String actualSourceLabel,
            String aladdinSourceLabel,
            String dispatchSourceLabel,
            /** 読込時の警告（フォールバック・部分失敗など）。空なら問題なし。 */
            String loadNotice,
            LoadStats loadStats) {

        public LoadedSources(
                ActualsSnapshot actuals,
                AladdinSnapshot aladdin,
                DispatchSnapshot dispatch,
                String actualSourceLabel,
                String aladdinSourceLabel,
                String dispatchSourceLabel) {
            this(
                    actuals,
                    aladdin,
                    dispatch,
                    actualSourceLabel,
                    aladdinSourceLabel,
                    dispatchSourceLabel,
                    "",
                    LoadStats.empty());
        }

        public LoadedSources(
                ActualsSnapshot actuals,
                AladdinSnapshot aladdin,
                DispatchSnapshot dispatch,
                String actualSourceLabel,
                String aladdinSourceLabel,
                String dispatchSourceLabel,
                String loadNotice) {
            this(
                    actuals,
                    aladdin,
                    dispatch,
                    actualSourceLabel,
                    aladdinSourceLabel,
                    dispatchSourceLabel,
                    loadNotice,
                    LoadStats.empty());
        }
    }

    private static final class SourceFileSizes {
        long actualBytes;
        long aladdinBytes;
        long dispatchBytes;

        long total() {
            return actualBytes + aladdinBytes + dispatchBytes;
        }
    }

    private EquipmentStatusDashboardSourceLoader() {}

    public static LoadedSources load(Map<String, String> ui) {
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
            Map<String, String> ui, SourceFingerprint previous, boolean haveCachedData) {
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

    private static LoadedSources loadSources(Map<String, String> env) {
        long t0 = System.nanoTime();
        SourceFileSizes sizes = resolveSourceFileSizes(env);
        StringBuilder notice = new StringBuilder();
        ActualsSnapshot actuals = loadActualsResilient(env, notice);
        AladdinSnapshot aladdin = loadAladdinResilient(env, notice);
        DispatchSnapshot dispatch = loadDispatch(env);
        long loadMs = Math.max(0L, (System.nanoTime() - t0) / 1_000_000L);
        LoadStats stats =
                new LoadStats(
                        sizes.total(),
                        loadMs,
                        rowCount(actuals),
                        rowCount(aladdin),
                        rowCount(dispatch));
        return new LoadedSources(
                actuals,
                aladdin,
                dispatch,
                actualsLabel(env),
                aladdinLabel(env),
                dispatchLabel(env),
                notice.toString().strip(),
                stats);
    }

    /** UI 向け: データサイズ・読込時間・行数の要約。 */
    public static String formatLoadStatsSummary(LoadStats stats) {
        if (stats == null) {
            return "";
        }
        NumberFormat nf = NumberFormat.getIntegerInstance(Locale.JAPAN);
        return "データ "
                + formatByteSize(stats.totalSourceBytes())
                + "  読込 "
                + formatLoadDuration(stats.loadDurationMs())
                + "  (行 実績 "
                + nf.format(stats.actualRowCount())
                + " / アラジン "
                + nf.format(stats.aladdinRowCount())
                + " / 配台 "
                + nf.format(stats.dispatchRowCount())
                + ")";
    }

    public static String formatByteSize(long bytes) {
        if (bytes < 0L) {
            bytes = 0L;
        }
        if (bytes < 1024L) {
            return bytes + " B";
        }
        if (bytes < 1024L * 1024L) {
            return String.format(Locale.ROOT, "%.1f KiB", bytes / 1024.0);
        }
        if (bytes < 1024L * 1024L * 1024L) {
            return String.format(Locale.ROOT, "%.1f MiB", bytes / (1024.0 * 1024.0));
        }
        return String.format(Locale.ROOT, "%.2f GiB", bytes / (1024.0 * 1024.0 * 1024.0));
    }

    public static String formatLoadDuration(long ms) {
        if (ms < 0L) {
            ms = 0L;
        }
        if (ms < 1000L) {
            return ms + " ms";
        }
        if (ms < 60_000L) {
            return String.format(Locale.ROOT, "%.2f s", ms / 1000.0);
        }
        long min = ms / 60_000L;
        long sec = (ms % 60_000L) / 1000L;
        return min + "分" + sec + "秒";
    }

    private static SourceFileSizes resolveSourceFileSizes(Map<String, String> env) {
        SourceFileSizes sizes = new SourceFileSizes();
        NetworkSourceDirResolver.Result r = NetworkSourceDirResolver.resolve(env);
        r.actualDetailPath().ifPresent(p -> sizes.actualBytes = fileSize(p));
        Optional<Path> taskInput = r.taskInputPath();
        if (taskInput.isPresent()) {
            sizes.aladdinBytes = fileSize(taskInput.get());
        } else {
            Path shaped = AppPaths.resolveShapedAladdinPlanJsonPath(env);
            if (Files.isRegularFile(shaped)) {
                sizes.aladdinBytes = fileSize(shaped);
            }
        }
        Path dispatch = AppPaths.resolveResultDispatchTableJsonPath(env);
        if (Files.isRegularFile(dispatch)) {
            sizes.dispatchBytes = fileSize(dispatch);
        }
        return sizes;
    }

    private static int rowCount(ActualsSnapshot snapshot) {
        return snapshot != null && snapshot.rows() != null ? snapshot.rows().size() : 0;
    }

    private static int rowCount(AladdinSnapshot snapshot) {
        return snapshot != null && snapshot.rows() != null ? snapshot.rows().size() : 0;
    }

    private static int rowCount(DispatchSnapshot snapshot) {
        return snapshot != null && snapshot.rows() != null ? snapshot.rows().size() : 0;
    }

    private static ActualsSnapshot loadActualsResilient(Map<String, String> ui, StringBuilder notice) {
        NetworkSourceDirResolver.Result r = NetworkSourceDirResolver.resolve(ui);
        Optional<Path> resolved = r.actualDetailPath();
        if (resolved.isPresent()) {
            Path file = resolved.get().toAbsolutePath().normalize();
            Optional<NetworkSourceFileReloadCache.Snapshot> mem =
                    NetworkSourceFileReloadCache.matchActuals(file);
            if (mem.isPresent()) {
                return EquipmentStatusDashboardBuilder.actualsFrom(mem.get().toTabularSheet());
            }
            Optional<ActualsSnapshot> shaped = loadActualsFromShapedJson(ui);
            long poiMax = dashboardActualsPoiMaxBytes(ui);
            long size = fileSize(file);
            if (shaped.isPresent() && (size > poiMax || isShapedActualsJsonFresh(file, ui))) {
                ActualsSnapshot snap = shaped.get();
                NetworkSourceFileReloadCache.storeActuals(
                        file,
                        isExcelPath(file),
                        List.of(),
                        0,
                        new PlanInputTabularIo.TabularSheet(snap.headers(), snap.rows()));
                return snap;
            }
            if (size > poiMax) {
                appendNotice(
                        notice,
                        "実績: ファイルが大きいため Excel 直読込を省略（"
                                + size
                                + " バイト）。加工実績DATAタブで一度読込し shaped_processing_actuals.json を生成してください");
                return shaped.orElse(new ActualsSnapshot(List.of(), List.of()));
            }
            try {
                return loadActualsFromResolvedFile(ui);
            } catch (Throwable ex) {
                appendNotice(notice, "実績: " + shortError(ex));
                return shaped.orElse(new ActualsSnapshot(List.of(), List.of()));
            }
        }
        return loadActualsFromShapedJson(ui).orElse(new ActualsSnapshot(List.of(), List.of()));
    }

    /** ダッシュボードで POI 直読込する実績の上限（加工実績タブ上限が 0 のときも既定 20MiB）。 */
    private static long dashboardActualsPoiMaxBytes(Map<String, String> ui) {
        long max = AppPaths.resolveActualDetailRawMaxBytes(ui);
        return max > 0 ? max : AppPaths.DEFAULT_PM_AI_ACTUAL_DETAIL_RAW_MAX_BYTES;
    }

    private static boolean isShapedActualsJsonFresh(Path sourceFile, Map<String, String> ui) {
        try {
            Path shaped = AppPaths.resolveShapedProcessingActualsJsonPath(ui);
            if (!Files.isRegularFile(shaped) || !Files.isRegularFile(sourceFile)) {
                return false;
            }
            long srcMod = Files.getLastModifiedTime(sourceFile).toMillis();
            long jsonMod = Files.getLastModifiedTime(shaped).toMillis();
            return jsonMod >= srcMod - 60_000L;
        } catch (IOException ex) {
            return false;
        }
    }

    private static long fileSize(Path file) {
        try {
            return Files.isRegularFile(file) ? Files.size(file) : 0L;
        } catch (IOException ex) {
            return 0L;
        }
    }

    private static AladdinSnapshot loadAladdinResilient(Map<String, String> ui, StringBuilder notice) {
        NetworkSourceDirResolver.Result r = NetworkSourceDirResolver.resolve(ui);
        Optional<Path> resolved = r.taskInputPath();
        if (resolved.isPresent()) {
            try {
                return loadAladdinFromFile(resolved.get());
            } catch (Throwable ex) {
                appendNotice(notice, "アラジン: " + shortError(ex));
            }
        }
        return loadAladdinFromShapedJson(ui).orElse(new AladdinSnapshot(List.of(), List.of()));
    }

    private static void appendNotice(StringBuilder notice, String line) {
        if (line == null || line.isBlank()) {
            return;
        }
        if (!notice.isEmpty()) {
            notice.append(' ');
        }
        notice.append(line.strip());
    }

    private static String shortError(Throwable ex) {
        if (ex instanceof OutOfMemoryError) {
            return "メモリ不足（heap space）。メモリ設定タブで JVM ヒープを増やすか、shaped JSON を利用してください";
        }
        String msg = ex.getMessage();
        if (msg == null || msg.isBlank()) {
            return ex.getClass().getSimpleName();
        }
        return msg.length() > 120 ? msg.substring(0, 120) + "…" : msg;
    }

    private static Optional<ActualsSnapshot> loadActualsFromShapedJson(Map<String, String> ui) {
        Path path = AppPaths.resolveShapedProcessingActualsJsonPath(ui);
        if (!Files.isRegularFile(path)) {
            return Optional.empty();
        }
        try {
            JsonTableIo.ArrayTable t = JsonTableIo.loadArrayTable(path);
            if (t.columns().isEmpty() && t.rows().isEmpty()) {
                return Optional.empty();
            }
            return Optional.of(
                    EquipmentStatusDashboardBuilder.normalizeActualsSnapshot(
                            new ActualsSnapshot(t.columns(), t.rows())));
        } catch (Exception ex) {
            return Optional.empty();
        }
    }

    private static Optional<AladdinSnapshot> loadAladdinFromShapedJson(Map<String, String> ui) {
        Path path = AppPaths.resolveShapedAladdinPlanJsonPath(ui);
        if (!Files.isRegularFile(path)) {
            return Optional.empty();
        }
        try {
            JsonTableIo.ArrayTable t = JsonTableIo.loadArrayTable(path);
            if (t.columns().isEmpty() && t.rows().isEmpty()) {
                return Optional.empty();
            }
            return Optional.of(new AladdinSnapshot(t.columns(), t.rows()));
        } catch (Exception ex) {
            return Optional.empty();
        }
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

    private static ActualsSnapshot loadActualsFromResolvedFile(Map<String, String> ui) throws IOException {
        NetworkSourceDirResolver.Result r = NetworkSourceDirResolver.resolve(ui);
        Optional<Path> resolved = r.actualDetailPath();
        if (resolved.isEmpty()) {
            return new ActualsSnapshot(List.of(), List.of());
        }
        Path file = resolved.get().toAbsolutePath().normalize();
        // matchActuals は loadActualsResilient 側で先に試す
        String low = file.getFileName().toString().toLowerCase(Locale.ROOT);
        if (low.endsWith(".pq") || low.endsWith(".parquet")) {
            return new ActualsSnapshot(List.of(), List.of());
        }
        int sheetIdx = 0;
        List<String> names = List.of();
        if (isExcelPath(file)) {
            names = TaskInputSourceRawGridIo.listExcelSheetNames(file);
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
        NetworkSourceFileReloadCache.storeActuals(file, isExcelPath(file), names, sheetIdx, shaped);
        saveShapedActualsJsonCache(ui, shaped);
        return EquipmentStatusDashboardBuilder.actualsFrom(shaped);
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
        NetworkSourceFileReloadCache.storeAladdin(
                normalized, isExcelPath(normalized), List.of(), 0, shaped);
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
        return r.actualDetailPath()
                .map(p -> p.getFileName().toString())
                .orElseGet(
                        () -> {
                            Path shaped = AppPaths.resolveShapedProcessingActualsJsonPath(ui);
                            if (Files.isRegularFile(shaped)) {
                                return shaped.getFileName().toString() + " (キャッシュJSON)";
                            }
                            return "(未設定)";
                        });
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

    /** 読込エラー表示・ログ向けに、解決済みソースパスを1ブロックで返す。 */
    public static String formatSourceContext(Map<String, String> ui) {
        Map<String, String> env = ui != null ? ui : Map.of();
        NetworkSourceDirResolver.Result r = NetworkSourceDirResolver.resolve(env);
        String actual =
                r.actualDetailPath()
                        .map(p -> p.toAbsolutePath().normalize().toString())
                        .orElse("(未設定 — PM_AI_ACTUAL_DETAIL_SOURCE_DIR / WORKBOOK を確認)");
        String aladdin =
                r.taskInputPath()
                        .map(p -> p.toAbsolutePath().normalize().toString())
                        .orElseGet(
                                () -> {
                                    Path shaped = AppPaths.resolveShapedAladdinPlanJsonPath(env);
                                    if (Files.isRegularFile(shaped)) {
                                        return shaped.toAbsolutePath().normalize()
                                                + " (shaped_aladdin JSON)";
                                    }
                                    return "(なし — PM_AI_TASK_INPUT_SOURCE_DIR 等を確認)";
                                });
        Path dispatch = AppPaths.resolveResultDispatchTableJsonPath(env);
        String dispatchPath =
                Files.isRegularFile(dispatch)
                        ? dispatch.toAbsolutePath().normalize().toString()
                        : "(なし — 結果_配台表.json を確認)";
        String sheet = env.getOrDefault(AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SHEET, "").strip();
        StringBuilder sb = new StringBuilder();
        sb.append("実績: ").append(actual);
        if (!sheet.isEmpty()) {
            sb.append("\n実績シート: ").append(sheet);
        }
        sb.append("\nアラジン: ").append(aladdin);
        sb.append("\n配台: ").append(dispatchPath);
        return sb.toString();
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

    /**
     * キャッシュ用 JSON に含める実績の主要見出し。
     * 加工実績DATAタブ・設備稼働ダッシュボード・加工トレンド・納期管理ビューで参照される全列を包含する。
     */
    public static final Set<String> ESSENTIAL_ACTUALS_HEADERS = Set.of(
            "工程名",
            "機械名",
            "依頼NO",
            "依頼ＮＯ",
            "加工日",
            "実加工数",
            "製造条件(内訳)",
            "加工開始日時",
            "加工終了日時",
            "換算数量",
            "累積実績",
            "累積完了率",
            "メンバー名");

    /**
     * 整形済み表から {@link #ESSENTIAL_ACTUALS_HEADERS} に合致する列を射影したキャッシュ用シートを生成する。
     */
    static PlanInputTabularIo.TabularSheet projectShapedForCache(PlanInputTabularIo.TabularSheet shaped) {
        if (shaped == null || shaped.headers().isEmpty()) {
            return shaped;
        }
        List<String> shHeaders = shaped.headers();
        List<Integer> keptIndices = new ArrayList<>();
        List<String> outHeaders = new ArrayList<>();
        for (int i = 0; i < shHeaders.size(); i++) {
            String title = shHeaders.get(i) != null ? shHeaders.get(i).strip() : "";
            if (ESSENTIAL_ACTUALS_HEADERS.contains(title)) {
                keptIndices.add(i);
                outHeaders.add(shHeaders.get(i));
            }
        }
        if (outHeaders.isEmpty()) {
            return shaped;
        }
        List<List<String>> outRows = new ArrayList<>(shaped.rows().size());
        for (List<String> row : shaped.rows()) {
            List<String> line = new ArrayList<>(keptIndices.size());
            for (int ix : keptIndices) {
                line.add(ix < row.size() && row.get(ix) != null ? row.get(ix) : "");
            }
            outRows.add(line);
        }
        return new PlanInputTabularIo.TabularSheet(outHeaders, outRows);
    }

    private static void saveShapedActualsJsonCache(
            Map<String, String> ui, PlanInputTabularIo.TabularSheet shaped) {
        if (shaped == null || shaped.headers().isEmpty() || shaped.rows().isEmpty()) {
            return;
        }
        try {
            Path savePath = AppPaths.resolveShapedProcessingActualsJsonPath(ui);
            PlanInputTabularIo.TabularSheet projected = projectShapedForCache(shaped);
            JsonTableIo.saveArrayTable(
                    savePath,
                    projected.headers(),
                    projected.rows(),
                    shaped.headers());
        } catch (Exception ex) {
            // キャッシュ保存失敗時も主フローを阻害しない
        }
    }
}
