package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.io.PoiWorkbookOpener;

/** 依頼書原本フォルダ内の各 xlsm「目次」シート行を一覧化する（閲覧専用・読み取りのみ）。 */
public final class RequestFormOriginalIndexSheetCatalog {

    private RequestFormOriginalIndexSheetCatalog() {}

    public record Row(
            String sourceFileName,
            String sourcePath,
            String iraiNo,
            String orderRequestDate,
            String responseDate,
            String inputDate,
            String deliveryDate,
            String deliveryRemarks,
            String contractDate,
            String contractNo,
            String contractRemarks) {}

    @FunctionalInterface
    public interface ScanProgressListener {
        void onProgress(int processed, int total);
    }

    public static List<Row> loadAll(Map<String, String> ui, List<String> warnings) {
        return loadAll(ui, warnings, null);
    }

    public static List<Row> loadAll(
            Map<String, String> ui, List<String> warnings, ScanProgressListener progress) {
        List<Row> out = new ArrayList<>();
        Path dir = AppPaths.resolveRequestFormOriginalDir(ui);
        if (!NetworkSourceDirResolver.isRequestFormOriginalDirReachable(ui)) {
            warn(warnings, "依頼書原本フォルダにアクセスできません: " + dir);
            return List.of();
        }
        File[] files =
                dir.toFile()
                        .listFiles(
                                (d, name) ->
                                        name.endsWith(".xlsm")
                                                && !name.startsWith("~$")
                                                && !name.equals("加工依頼書入力.xlsm"));
        if (files == null || files.length == 0) {
            warn(warnings, "Excel 依頼書原本が見つかりません: " + dir);
            return List.of();
        }
        int processed = 0;
        for (File file : files) {
            try (Workbook wb = PoiWorkbookOpener.open(file)) {
                Sheet index = wb.getSheet(RequestFormOriginalIndexSheetLayout.SHEET_NAME);
                if (index == null) {
                    continue;
                }
                String pathText = file.toPath().toAbsolutePath().normalize().toString();
                String fileName = file.getName();
                for (RequestFormOriginalIndexSheetReader.IndexEntry entry :
                        RequestFormOriginalIndexSheetReader.read(index).values()) {
                    out.add(toRow(fileName, pathText, entry));
                }
            } catch (Exception ex) {
                warn(warnings, "目次シート読込エラー " + file.getName() + ": " + ex.getMessage());
            }
            processed++;
            if (progress != null) {
                progress.onProgress(processed, files.length);
            }
        }
        out.sort(
                Comparator.comparing(Row::sourceFileName)
                        .thenComparing(Row::iraiNo, String.CASE_INSENSITIVE_ORDER));
        return List.copyOf(out);
    }

    private static Row toRow(
            String fileName, String pathText, RequestFormOriginalIndexSheetReader.IndexEntry entry) {
        return new Row(
                fileName,
                pathText,
                entry.iraiNo(),
                entry.orderRequestDate(),
                entry.responseDate(),
                entry.inputDate(),
                entry.deliveryDate(),
                entry.deliveryRemarks(),
                entry.contractDate(),
                entry.contractNo(),
                entry.contractRemarks());
    }

    private static void warn(List<String> warnings, String message) {
        if (warnings != null) {
            warnings.add(message);
        }
    }
}
