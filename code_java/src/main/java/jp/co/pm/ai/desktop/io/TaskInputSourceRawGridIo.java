package jp.co.pm.ai.desktop.io;

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.DataFormatter;

/**
 * {@link jp.co.pm.ai.desktop.config.AppPaths#KEY_PM_AI_TASK_INPUT_SOURCE_DIR} 由来ブックの「生シート」表示用。先頭行をヘッダー専用とせず、
 * シート上の全行をデータとして読み、列見出しは {@code 列1..列N} とする。
 */
public final class TaskInputSourceRawGridIo {

    private static final DataFormatter CELL_FORMAT = new DataFormatter();

    private TaskInputSourceRawGridIo() {}

    /**
     * タスク入力ディレクトリで選ばれたファイルを生表として読む。
     *
     * @param excelSheetIndex Excel 系のみ有効（0 から）。CSV は無視される。
     */
    public static PlanInputTabularIo.TabularSheet readRaw(Path path, int excelSheetIndex)
            throws IOException {
        String low = path.getFileName().toString().toLowerCase(Locale.ROOT);
        if (low.endsWith(".csv")) {
            return readCsvRaw(path);
        }
        if (low.endsWith(".xlsx")
                || low.endsWith(".xlsm")
                || low.endsWith(".xltx")
                || low.endsWith(".xltm")) {
            return readExcelSheetRaw(path, excelSheetIndex);
        }
        throw new IOException("未対応の拡張子（csv / xlsx / xlsm / xltx / xltm のみ）: " + path);
    }

    /** Excel ブックのシート名一覧（順序はブック定義順）。 */
    public static List<String> listExcelSheetNames(Path path) throws IOException {
        String low = path.getFileName().toString().toLowerCase(Locale.ROOT);
        if (!(low.endsWith(".xlsx")
                || low.endsWith(".xlsm")
                || low.endsWith(".xltx")
                || low.endsWith(".xltm"))) {
            return List.of();
        }
        try (Workbook wb = WorkbookFactory.create(path.toFile())) {
            List<String> names = new ArrayList<>();
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                names.add(wb.getSheetName(i));
            }
            return names;
        }
    }

    private static PlanInputTabularIo.TabularSheet readCsvRaw(Path path) throws IOException {
        List<List<String>> allRows = new ArrayList<>();
        int maxCol = 0;
        try (BufferedReader r = Files.newBufferedReader(path, StandardCharsets.UTF_8)) {
            String line;
            while ((line = r.readLine()) != null) {
                List<String> cells = new ArrayList<>();
                PlanInputTabularIo.parseCsvLine(line, cells);
                maxCol = Math.max(maxCol, cells.size());
                allRows.add(cells);
            }
        }
        List<String> headers = new ArrayList<>();
        for (int c = 0; c < maxCol; c++) {
            headers.add("列" + (c + 1));
        }
        List<List<String>> rows = new ArrayList<>();
        for (List<String> src : allRows) {
            List<String> line = new ArrayList<>(maxCol);
            for (int c = 0; c < maxCol; c++) {
                line.add(c < src.size() && src.get(c) != null ? src.get(c) : "");
            }
            rows.add(line);
        }
        return new PlanInputTabularIo.TabularSheet(headers, rows);
    }

    private static PlanInputTabularIo.TabularSheet readExcelSheetRaw(Path path, int sheetIndex)
            throws IOException {
        try (Workbook wb = WorkbookFactory.create(path.toFile())) {
            if (sheetIndex < 0 || sheetIndex >= wb.getNumberOfSheets()) {
                throw new IOException(
                        "sheet index out of range: " + sheetIndex + " (sheets=" + wb.getNumberOfSheets() + ")");
            }
            Sheet sh = wb.getSheetAt(sheetIndex);
            int lastRow = sh.getLastRowNum();
            int maxCol = 0;
            for (int r = 0; r <= lastRow; r++) {
                Row row = sh.getRow(r);
                if (row != null) {
                    maxCol = Math.max(maxCol, row.getLastCellNum());
                }
            }
            if (maxCol < 0) {
                maxCol = 0;
            }
            List<String> headers = new ArrayList<>();
            for (int c = 0; c < maxCol; c++) {
                headers.add("列" + (c + 1));
            }
            List<List<String>> rows = new ArrayList<>();
            for (int r = 0; r <= lastRow; r++) {
                Row row = sh.getRow(r);
                List<String> line = new ArrayList<>(maxCol);
                for (int c = 0; c < maxCol; c++) {
                    line.add(row == null ? "" : cellToString(row.getCell(c)));
                }
                rows.add(line);
            }
            return new PlanInputTabularIo.TabularSheet(headers, rows);
        }
    }

    private static String cellToString(Cell cell) {
        if (cell == null) {
            return "";
        }
        return CELL_FORMAT.formatCellValue(cell);
    }
}
