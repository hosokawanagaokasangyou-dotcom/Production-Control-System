package jp.co.pm.ai.desktop.reconciliation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.Map;

/** 依頼書原本 xlsm の「目次」シートから依頼No 別の行データを読む。 */
final class RequestFormOriginalIndexSheetReader {

    /** 目次 1 行分（rawMap 未反映列も将来用に保持）。 */
    record IndexEntry(
            String iraiNo,
            String orderRequestDate,
            String responseDate,
            String inputDate,
            String deliveryDate,
            String deliveryRemarks,
            String contractDate,
            String contractNo,
            String contractRemarks) {}

    private static final DataFormatter CELL_FORMATTER = new DataFormatter();

    private RequestFormOriginalIndexSheetReader() {}

    static Map<String, IndexEntry> read(Sheet sheet) {
        if (sheet == null) {
            return Map.of();
        }
        int headerRow = findHeaderRowIndex(sheet);
        if (headerRow < 0) {
            return Map.of();
        }
        FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
        LinkedHashMap<String, IndexEntry> out = new LinkedHashMap<>();
        int lastRow = sheet.getLastRowNum();
        for (int rowIndex = headerRow + 1; rowIndex <= lastRow; rowIndex++) {
            String iraiNo = cellString(sheet, rowIndex, RequestFormOriginalIndexSheetLayout.COL_IRAI_NO, evaluator);
            if (!isDataIraiNo(iraiNo)) {
                continue;
            }
            IndexEntry entry =
                    new IndexEntry(
                            iraiNo.strip(),
                            cellString(
                                    sheet,
                                    rowIndex,
                                    RequestFormOriginalIndexSheetLayout.COL_ORDER_REQUEST_DATE,
                                    evaluator),
                            cellString(
                                    sheet,
                                    rowIndex,
                                    RequestFormOriginalIndexSheetLayout.COL_RESPONSE_DATE,
                                    evaluator),
                            cellString(
                                    sheet,
                                    rowIndex,
                                    RequestFormOriginalIndexSheetLayout.COL_INPUT_DATE,
                                    evaluator),
                            cellString(
                                    sheet,
                                    rowIndex,
                                    RequestFormOriginalIndexSheetLayout.COL_DELIVERY_DATE,
                                    evaluator),
                            cellString(
                                    sheet,
                                    rowIndex,
                                    RequestFormOriginalIndexSheetLayout.COL_DELIVERY_REMARKS,
                                    evaluator),
                            cellString(
                                    sheet,
                                    rowIndex,
                                    RequestFormOriginalIndexSheetLayout.COL_CONTRACT_DATE,
                                    evaluator),
                            cellString(
                                    sheet,
                                    rowIndex,
                                    RequestFormOriginalIndexSheetLayout.COL_CONTRACT_NO,
                                    evaluator),
                            cellString(
                                    sheet,
                                    rowIndex,
                                    RequestFormOriginalIndexSheetLayout.COL_CONTRACT_REMARKS,
                                    evaluator));
            String key = JuchuTransferValueNormalizer.normalizeKey(entry.iraiNo());
            if (!key.isEmpty()) {
                out.put(key, entry);
            }
        }
        return Map.copyOf(out);
    }

    private static int findHeaderRowIndex(Sheet sheet) {
        int limit =
                Math.min(
                        sheet.getLastRowNum(),
                        RequestFormOriginalIndexSheetLayout.HEADER_SCAN_MAX_ROW);
        for (int rowIndex = 0; rowIndex <= limit; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }
            Cell cell = row.getCell(RequestFormOriginalIndexSheetLayout.COL_IRAI_NO);
            if (cell == null) {
                continue;
            }
            String text = cell.getStringCellValue();
            if (text != null && text.contains("加工依頼")) {
                return rowIndex;
            }
        }
        return -1;
    }

    private static boolean isDataIraiNo(String iraiNo) {
        if (iraiNo == null || iraiNo.isBlank()) {
            return false;
        }
        String text = iraiNo.strip();
        if (text.contains("加工依頼")) {
            return false;
        }
        return true;
    }

    private static String cellString(
            Sheet sheet, int rowIndex, int colIndex, FormulaEvaluator evaluator) {
        Row row = sheet.getRow(rowIndex);
        Cell cell = row != null ? row.getCell(colIndex) : null;
        if (cell == null) {
            return "";
        }
        try {
            return CELL_FORMATTER.formatCellValue(cell, evaluator).strip();
        } catch (Exception ex) {
            return "";
        }
    }
}
