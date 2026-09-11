package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.TimeoutException;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;

/** 検査表索引の更新と Excel 起動。 */
public final class InspectionSheetOpenService {

    public record RebuildResult(
            List<InspectionSheetIndexStore.Row> rows, List<String> warnings, int readExcelCount) {}

    private static final Object IN_FLIGHT_LOCK = new Object();
    private static String inFlightKey;
    private static CompletableFuture<RebuildResult> inFlight;

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
        if (!dirReachable(snap)) {
            Path dir = resolveDir(snap);
            return CompletableFuture.completedFuture(
                    new RebuildResult(List.of(), List.of("検査表フォルダにアクセスできません: " + dir), 0));
        }
        String key = rebuildKey(snap);
        synchronized (IN_FLIGHT_LOCK) {
            if (inFlight != null && !inFlight.isDone() && key.equals(inFlightKey)) {
                return inFlight;
            }
            CompletableFuture<RebuildResult> future = new CompletableFuture<>();
            inFlight = future;
            inFlightKey = key;
            Thread t =
                    new Thread(
                            () -> {
                                try {
                                    future.complete(rebuild(snap, progress));
                                } catch (Throwable ex) {
                                    future.completeExceptionally(ex);
                                }
                            },
                            "inspection-sheet-index-warmup");
            t.setDaemon(true);
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
        Path dir = resolveDir(ui);
        FactorySite site = factorySite(ui);
        Path csv = InspectionSheetIndexStore.indexFile(site);
        List<InspectionSheetIndexStore.Row> previous = InspectionSheetIndexStore.load(csv);
        InspectionSheetIndexScanner.Result scanned = InspectionSheetIndexScanner.scan(dir, previous, progress);
        InspectionSheetIndexStore.save(csv, scanned.rows());
        return new RebuildResult(scanned.rows(), scanned.warnings(), scanned.readExcelCount());
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
        RebuildResult rebuilt = joinRebuild(ui);
        return InspectionSheetLookup.find(rebuilt.rows(), iraiNo);
    }

    public static void open(InspectionSheetIndexStore.Row row) throws IOException {
        if (row == null || row.filePath() == null || row.filePath().isBlank()) {
            throw new IOException("検査表ファイルパスが空です");
        }
        DesktopFileOpener.openFile(Path.of(row.filePath()));
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
        Path dir = resolveDir(ui);
        String siteKey = site != null ? site.name() : "";
        String dirKey = dir != null ? dir.toAbsolutePath().normalize().toString() : "";
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
}
