package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.io.DesktopFileOpener;

/** 検査表索引の更新と Excel 起動。 */
public final class InspectionSheetOpenService {

    public record RebuildResult(
            List<InspectionSheetIndexStore.Row> rows, List<String> warnings, int readExcelCount) {}

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
        if (rows.isEmpty()) {
            RebuildResult rebuilt = rebuild(ui, null);
            return InspectionSheetLookup.find(rebuilt.rows(), iraiNo);
        }
        return hits;
    }

    public static void open(InspectionSheetIndexStore.Row row) throws IOException {
        if (row == null || row.filePath() == null || row.filePath().isBlank()) {
            throw new IOException("検査表ファイルパスが空です");
        }
        DesktopFileOpener.openFile(Path.of(row.filePath()));
    }
}
