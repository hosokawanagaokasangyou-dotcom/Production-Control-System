package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.PoiWorkbookOpener;
import jp.co.pm.ai.desktop.reconciliation.JuchuSheetColumnLayout.Col;

/**
 * 依頼NO → EC面区分（両面EC/片面EC）の lookup。
 *
 * <p>優先: 配台計画_タスク入力 xlsx の {@link EcSideClassification#COLUMN_TITLE} 列。
 * フォールバック: 受注ファイルから再判定。
 */
public final class EcSideClassificationLookup {

    private static final String JUCHU_SHEET_NAME = "受注ﾌｧｲﾙ";
    private static final int JUCHU_SHEET_MAX_SCAN_ROWS = 20_000;
    private static final String COL_TID = "依頼NO";

    private EcSideClassificationLookup() {}

    public static Map<String, String> loadByIraiNoKey(Map<String, String> ui, List<String> warnings) {
        Map<String, String> out = new LinkedHashMap<>();
        loadFromPlanInput(ui, out, warnings);
        loadMissingFromJuchu(ui, out, warnings);
        return Map.copyOf(out);
    }

    private static void loadFromPlanInput(
            Map<String, String> ui, Map<String, String> out, List<String> warnings) {
        Path planPath = resolvePlanInputPath(ui);
        if (planPath == null || !Files.isRegularFile(planPath)) {
            return;
        }
        try {
            PlanInputTabularIo.TabularRead read =
                    PlanInputTabularIo.readWithResolvedSheet(
                            planPath, AppPaths.STAGE1_PLAN_OUTPUT_SHEET);
            List<String> headers = read.tabular().headers();
            int tidIdx = indexOfHeader(headers, COL_TID);
            int ecIdx = indexOfHeader(headers, EcSideClassification.COLUMN_TITLE);
            if (tidIdx < 0 || ecIdx < 0) {
                return;
            }
            for (List<String> row : read.tabular().rows()) {
                if (row == null || tidIdx >= row.size() || ecIdx >= row.size()) {
                    continue;
                }
                String tid = row.get(tidIdx) != null ? row.get(tidIdx).strip() : "";
                String ecClass = row.get(ecIdx) != null ? row.get(ecIdx).strip() : "";
                if (tid.isEmpty() || ecClass.isEmpty()) {
                    continue;
                }
                if (!EcSideClassification.DOUBLE_SIDED.equals(ecClass)
                        && !EcSideClassification.SINGLE_SIDED.equals(ecClass)) {
                    continue;
                }
                out.putIfAbsent(RequestFormOriginalIndexLookup.normalizeIraiNoKey(tid), ecClass);
            }
        } catch (Exception ex) {
            warn(warnings, "配台計画タスク入力の EC面区分読込エラー: " + ex.getMessage());
        }
    }

    private static void loadMissingFromJuchu(
            Map<String, String> ui, Map<String, String> out, List<String> warnings) {
        String juchuPath =
                AppPaths.resolveRequestFormJuchuFile(ui).map(Path::toString).orElse("");
        if (juchuPath.isBlank()) {
            return;
        }
        File juchuFile = new File(juchuPath);
        if (!juchuFile.isFile()) {
            return;
        }
        JuchuHeaderAliasRegistry registry = JuchuHeaderAliasRegistry.loadDefault();
        try (FileInputStream fis = new FileInputStream(juchuFile);
                Workbook wb = PoiWorkbookOpener.open(fis)) {
            Sheet sheet = wb.getSheet(JUCHU_SHEET_NAME);
            if (sheet == null) {
                return;
            }
            int firstDataRow = registry.headerRowIndexFor(juchuPath) + 1;
            int reqNoColIdx =
                    JuchuSheetColumnLayout.resolveTransferColumnIndex(Col.IRAI_NO, registry, juchuPath);
            int lastDataRow = findLastPopulatedDataRow(sheet, firstDataRow, reqNoColIdx);
            for (int r = firstDataRow; r <= lastDataRow; r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    continue;
                }
                String reqNo = readCellAsString(row.getCell(reqNoColIdx)).strip();
                if (reqNo.isEmpty()) {
                    continue;
                }
                String normKey = RequestFormOriginalIndexLookup.normalizeIraiNoKey(reqNo);
                if (out.containsKey(normKey)) {
                    continue;
                }
                Map<String, String> vals =
                        JuchuSheetColumnLayout.readDbValuesFromRow(row, registry, juchuPath);
                String kako = valueForKey(vals, Col.KAKO_NAIYO.dbKey());
                String ecMen = valueForKey(vals, Col.EC_MEN.dbKey());
                String classified = EcSideClassification.classify(kako, ecMen);
                if (!classified.isEmpty()) {
                    out.put(normKey, classified);
                }
            }
        } catch (Exception ex) {
            warn(warnings, "受注ファイル EC面区分フォールバック読込エラー: " + ex.getMessage());
        }
    }

    private static Path resolvePlanInputPath(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = u.get(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH);
        if (override != null && !override.strip().isEmpty()) {
            Path p = Path.of(override.strip()).toAbsolutePath().normalize();
            if (Files.isRegularFile(p)) {
                return p;
            }
        }
        return AppPaths.defaultStage1PlanTasksPath(u);
    }

    private static int indexOfHeader(List<String> headers, String title) {
        if (headers == null || title == null) {
            return -1;
        }
        for (int i = 0; i < headers.size(); i++) {
            String h = headers.get(i);
            if (h != null && title.equals(h.strip())) {
                return i;
            }
        }
        return -1;
    }

    private static String valueForKey(Map<String, String> map, String key) {
        if (map == null || key == null) {
            return "";
        }
        String v = map.get(key);
        return v != null ? v : "";
    }

    private static int findLastPopulatedDataRow(Sheet sheet, int firstDataRow, int reqNoColIdx) {
        int last = sheet.getLastRowNum();
        if (last > firstDataRow + JUCHU_SHEET_MAX_SCAN_ROWS) {
            last = firstDataRow + JUCHU_SHEET_MAX_SCAN_ROWS;
        }
        for (int r = last; r >= firstDataRow; r--) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            String reqNo = readCellAsString(row.getCell(reqNoColIdx)).strip();
            if (!reqNo.isEmpty()) {
                return r;
            }
        }
        return firstDataRow - 1;
    }

    private static void warn(List<String> warnings, String message) {
        if (warnings != null) {
            warnings.add(message);
        }
    }

    private static String readCellAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> {
                try {
                    yield cell.getStringCellValue();
                } catch (Exception ex) {
                    yield String.valueOf(cell.getNumericCellValue());
                }
            }
            default -> "";
        };
    }
}
