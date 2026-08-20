package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * JSON 格子を既存 master ブックの skills / need / speed / 組み合わせ表へ書き戻す。
 * 他シート・VBA は開いたブックのまま残す。
 */
public final class MasterDispatchSheetWorkbookExporter {

    private static final Pattern NUMERIC =
            Pattern.compile("^-?(?:0|[1-9]\\d*)(?:\\.\\d+)?$");

    private MasterDispatchSheetWorkbookExporter() {}

    public static void writeBack(
            Path workbookPath, MasterDispatchSheetsDocument document, Map<String, String> ui)
            throws IOException {
        Objects.requireNonNull(workbookPath, "workbookPath");
        Objects.requireNonNull(document, "document");
        Path target = workbookPath.toAbsolutePath().normalize();
        if (!Files.isRegularFile(target)) {
            throw new IOException("master ブックが見つかりません: " + target);
        }
        Map<String, String> env = ui != null ? ui : Map.of();
        try (Workbook wb = PoiWorkbookOpener.open(target.toFile())) {
            for (String key : MasterDispatchSheetsDocument.SHEET_KEYS) {
                MasterDispatchSheetsDocument.SheetGrid grid = document.sheet(key);
                String sheetName =
                        grid.sheetName() != null && !grid.sheetName().isBlank()
                                ? grid.sheetName()
                                : MasterDispatchSheetsDocument.defaultSheetName(key);
                Sheet sh = wb.getSheet(sheetName);
                if (sh == null) {
                    sh = wb.createSheet(sheetName);
                }
                writeGrid(sh, grid.rows());
            }
            PoiWorkbookFileWriter.writeReplacing(target, wb, env);
        }
    }

    static void writeGrid(Sheet sh, List<List<String>> rows) {
        List<List<String>> data = rows != null ? rows : List.of();
        int newRows = data.size();
        int newCols = 0;
        for (List<String> row : data) {
            if (row != null) {
                newCols = Math.max(newCols, row.size());
            }
        }
        for (int r = 0; r < newRows; r++) {
            Row row = sh.getRow(r);
            if (row == null) {
                row = sh.createRow(r);
            }
            List<String> cells = data.get(r);
            for (int c = 0; c < newCols; c++) {
                String v = cells != null && c < cells.size() ? cells.get(c) : "";
                writeCell(row, c, v);
            }
            short last = row.getLastCellNum();
            int lastCol = last > 0 ? last : 0;
            for (int c = newCols; c < lastCol; c++) {
                Cell extra = row.getCell(c);
                if (extra != null) {
                    row.removeCell(extra);
                }
            }
        }
        for (int r = sh.getLastRowNum(); r >= newRows; r--) {
            Row leftover = sh.getRow(r);
            if (leftover != null) {
                sh.removeRow(leftover);
            }
        }
    }

    private static void writeCell(Row row, int col, String value) {
        String v = value != null ? value : "";
        Cell cell = row.getCell(col);
        if (v.isEmpty()) {
            if (cell != null) {
                cell.setBlank();
            }
            return;
        }
        if (cell == null) {
            cell = row.createCell(col);
        }
        if (NUMERIC.matcher(v).matches()) {
            cell.setCellValue(Double.parseDouble(v));
        } else {
            cell.setCellValue(v);
        }
    }
}
