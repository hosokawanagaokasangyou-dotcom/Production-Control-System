package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.TimeoutException;
import java.util.concurrent.atomic.AtomicInteger;

import javafx.application.Platform;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;

/** 検査表索引の更新と Excel 起動。 */
public final class InspectionSheetOpenService {

    public record RebuildResult(
            List<InspectionSheetIndexStore.Row> rows,
            List<String> warnings,
            int readExcelCount,
            InspectionSheetIndexShare.PullResult sharePull) {
        public RebuildResult {
            rows = rows != null ? List.copyOf(rows) : List.of();
            warnings = warnings != null ? List.copyOf(warnings) : List.of();
            sharePull = sharePull != null ? sharePull : InspectionSheetIndexShare.PullResult.SKIPPED;
        }

        public RebuildResult(
                List<InspectionSheetIndexStore.Row> rows, List<String> warnings, int readExcelCount) {
            this(rows, warnings, readExcelCount, InspectionSheetIndexShare.PullResult.SKIPPED);
        }
    }

    public static final int PARTIAL_SAVE_EVERY_EXCEL_READS = 20;

    private static final Object IN_FLIGHT_LOCK = new Object();
    private static String inFlightKey;
    private static CompletableFuture<RebuildResult> inFlight;
    private static final CopyOnWriteArrayList<InspectionSheetIndexScanner.Progress> inFlightProgress =
            new CopyOnWriteArrayList<>();
    private static final Object CHECKPOINT_LOCK = new Object();
    private static Path inFlightCsv;
    private static List<InspectionSheetIndexStore.Row> inFlightRows;

    static {
        Runtime.getRuntime()
                .addShutdownHook(
                        new Thread(
                                InspectionSheetOpenService::flushInFlightCheckpoint,
                                "inspection-sheet-index-checkpoint"));
    }

    private InspectionSheetOpenService() {}

    public static Path resolveDir(Map<String, String> ui) {
        return AppPaths.resolveInspectionSheetDir(ui);
    }

    public static boolean dirReachable(Map<String, String> ui) {
        return NetworkSourceDirResolver.isInspectionSheetDirReachable(ui);
    }

    public static FactorySite factorySite(Map<String, String> ui) {
        return GlobalInitSettingTarget.loadEffective(ui != null ? ui : Map.of());
    }

    public static CompletableFuture<RebuildResult> startBackgroundRebuild(Map<String, String> ui) {
        return startBackgroundRebuild(ui, null);
    }

    public static CompletableFuture<RebuildResult> startBackgroundRebuild(
            Map<String, String> ui, InspectionSheetIndexScanner.Progress progress) {
        Map<String, String> snap = snapshotUi(ui);
        String key = rebuildKey(snap);
        synchronized (IN_FLIGHT_LOCK) {
            if (inFlight != null && !inFlight.isDone() && key.equals(inFlightKey)) {
                if (progress != null) {
                    inFlightProgress.add(progress);
                }
                return inFlight;
            }
            inFlightProgress.clear();
            if (progress != null) {
                inFlightProgress.add(progress);
            }
            CompletableFuture<RebuildResult> future = new CompletableFuture<>();
            inFlight = future;
            inFlightKey = key;
            Thread t =
                    new Thread(
                            () -> {
                                try {
                                    if (!dirReachable(snap)) {
                                        Path dir = resolveDir(snap);
                                        future.complete(
                                                new RebuildResult(
                                                        List.of(),
                                                        List.of(
                                                                "検査表フォルダにアクセスできません: "
                                                                        + dir),
                                                        0));
                                        return;
                                    }
                                    future.complete(
                                            rebuild(snap, InspectionSheetOpenService::fanOutProgress));
                                } catch (Throwable ex) {
                                    future.completeExceptionally(ex);
                                } finally {
                                    synchronized (IN_FLIGHT_LOCK) {
                                        inFlightProgress.clear();
                                    }
                                }
                            },
                            "inspection-sheet-index-warmup");
            t.setDaemon(true);
            t.setPriority(Thread.MIN_PRIORITY);
            t.start();
            return future;
        }
    }

    static void joinBackgroundRebuildForTest() {
        CompletableFuture<RebuildResult> future;
        synchronized (IN_FLIGHT_LOCK) {
            future = inFlight;
        }
        if (future == null || future.isDone()) {
            return;
        }
        try {
            future.get(5, TimeUnit.SECONDS);
        } catch (InterruptedException ex) {
            Thread.currentThread().interrupt();
        } catch (ExecutionException | TimeoutException ignored) {
            future.completeExceptionally(new IOException("test teardown"));
        }
    }

    public static RebuildResult rebuild(
            Map<String, String> ui, InspectionSheetIndexScanner.Progress progress) throws IOException {
        List<Path> dirs = AppPaths.resolveInspectionSheetDirs(ui);
        FactorySite site = factorySite(ui);
        Path csv = InspectionSheetIndexStore.indexFile(site);
        InspectionSheetIndexShare.PullResult sharePull = InspectionSheetIndexShare.PullResult.SKIPPED;
        try {
            sharePull = InspectionSheetIndexShare.pullIfNeeded(ui);
        } catch (IOException ignored) {
            // 共有が届かないときはローカルのまま増分更新する
        }
        List<InspectionSheetIndexStore.Row> previous = InspectionSheetIndexStore.loadMerged(csv);
        AtomicInteger lastSavedExcel = new AtomicInteger(0);
        beginCheckpoint(csv);
        try {
            InspectionSheetIndexScanner.Result scanned =
                    InspectionSheetIndexScanner.scan(
                            dirs,
                            previous,
                            progress,
                            (rows, excel) -> {
                                rememberCheckpoint(csv, rows);
                                if (excel > 0
                                        && excel - lastSavedExcel.get()
                                                >= PARTIAL_SAVE_EVERY_EXCEL_READS) {
                                    savePartialQuiet(csv, rows);
                                    lastSavedExcel.set(excel);
                                }
                            });
            InspectionSheetIndexStore.save(csv, scanned.rows());
            InspectionSheetIndexStore.clearPartial(csv);
            return new RebuildResult(
                    scanned.rows(), scanned.warnings(), scanned.readExcelCount(), sharePull);
        } catch (IOException | RuntimeException ex) {
            flushInFlightCheckpoint();
            throw ex;
        } finally {
            endCheckpoint(csv);
        }
    }

    public static void flushInFlightCheckpoint() {
        Path csv;
        List<InspectionSheetIndexStore.Row> rows;
        synchronized (CHECKPOINT_LOCK) {
            csv = inFlightCsv;
            rows = inFlightRows;
        }
        savePartialQuiet(csv, rows);
    }

    public static List<InspectionSheetIndexStore.Row> loadIndex(Map<String, String> ui) throws IOException {
        return InspectionSheetIndexStore.load(InspectionSheetIndexStore.indexFile(factorySite(ui)));
    }

    public static List<InspectionSheetIndexStore.Row> find(
            Map<String, String> ui, String iraiNo) throws IOException {
        List<InspectionSheetIndexStore.Row> rows = loadIndex(ui);
        List<InspectionSheetIndexStore.Row> hits = InspectionSheetLookup.find(rows, iraiNo);
        if (!hits.isEmpty()) {
            return hits;
        }
        if (!rows.isEmpty()) {
            return hits;
        }
        if (isJavaFxApplicationThread()) {
            startBackgroundRebuild(ui);
            return hits;
        }
        RebuildResult rebuilt = joinRebuild(ui);
        return InspectionSheetLookup.find(rebuilt.rows(), iraiNo);
    }

    public static void open(InspectionSheetIndexStore.Row row) throws IOException {
        if (row == null || row.filePath() == null || row.filePath().isBlank()) {
            throw new IOException("検査表ファイルパスが空です");
        }
        DesktopFileOpener.openFile(Path.of(row.filePath()));
    }

    static boolean isRebuildInFlight() {
        synchronized (IN_FLIGHT_LOCK) {
            return inFlight != null && !inFlight.isDone();
        }
    }

    private static boolean isJavaFxApplicationThread() {
        try {
            return Platform.isFxApplicationThread();
        } catch (IllegalStateException ignored) {
            return false;
        }
    }

    private static RebuildResult joinRebuild(Map<String, String> ui) throws IOException {
        try {
            return startBackgroundRebuild(ui).get();
        } catch (InterruptedException ex) {
            Thread.currentThread().interrupt();
            throw new IOException("検査表索引の更新が中断されました", ex);
        } catch (ExecutionException ex) {
            Throwable cause = ex.getCause() != null ? ex.getCause() : ex;
            if (cause instanceof IOException io) {
                throw io;
            }
            throw new IOException(cause.getMessage(), cause);
        }
    }

    private static String rebuildKey(Map<String, String> ui) {
        FactorySite site = factorySite(ui);
        String siteKey = site != null ? site.name() : "";
        StringBuilder dirKey = new StringBuilder();
        for (Path dir : AppPaths.resolveInspectionSheetDirs(ui)) {
            if (dirKey.length() > 0) {
                dirKey.append(";");
            }
            dirKey.append(dir.toAbsolutePath().normalize());
        }
        return siteKey + "|" + dirKey;
    }

    private static Map<String, String> snapshotUi(Map<String, String> ui) {
        if (ui == null || ui.isEmpty()) {
            return Map.of();
        }
        Map<String, String> copy = new LinkedHashMap<>();
        ui.forEach(
                (k, v) -> {
                    if (k != null && v != null) {
                        copy.put(k, v);
                    }
                });
        return Map.copyOf(copy);
    }

    private static void fanOutProgress(String phase, int processed, int total) {
        for (InspectionSheetIndexScanner.Progress listener : inFlightProgress) {
            listener.onProgress(phase, processed, total);
        }
    }

    private static void beginCheckpoint(Path csv) {
        synchronized (CHECKPOINT_LOCK) {
            inFlightCsv = csv;
            inFlightRows = List.of();
        }
    }

    private static void rememberCheckpoint(Path csv, List<InspectionSheetIndexStore.Row> rows) {
        synchronized (CHECKPOINT_LOCK) {
            if (csv == null || !csv.equals(inFlightCsv)) {
                return;
            }
            inFlightRows = rows != null ? rows : List.of();
        }
    }

    private static void endCheckpoint(Path csv) {
        synchronized (CHECKPOINT_LOCK) {
            if (csv != null && csv.equals(inFlightCsv)) {
                inFlightCsv = null;
                inFlightRows = null;
            }
        }
    }

    private static void savePartialQuiet(Path csv, List<InspectionSheetIndexStore.Row> rows) {
        if (csv == null || rows == null || rows.isEmpty()) {
            return;
        }
        try {
            InspectionSheetIndexStore.savePartial(csv, rows);
        } catch (IOException ignored) {
            // 途中保存の失敗で本走査を止めない
        }
    }
}
