package jp.co.pm.ai.desktop.io;

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.DataFormatter;

/**
 * Raw sheet reader for {@link jp.co.pm.ai.desktop.config.AppPaths#KEY_PM_AI_TASK_INPUT_SOURCE_DIR}.
 * All sheet rows are data; synthetic headers use column index labels (see {@code readRaw} output).
 */
public final class TaskInputSourceRawGridIo {

    private static final DataFormatter CELL_FORMAT = new DataFormatter();

    private TaskInputSourceRawGridIo() {}

    /**
     * Reads the selected file as a raw grid (CSV or Excel sheet by index).
     *
     * @param excelSheetIndex Excel sheet index (0-based); ignored for CSV
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
        throw new IOException("unsupported extension (csv / xlsx / xlsm / xltx / xltm only): " + path);
    }

    /** Lists Excel workbook sheet names (empty list if not Excel). */
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
            headers.add("\u5217" + (c + 1));
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
                headers.add("\u5217" + (c + 1));
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

    /**
     * Aladdin tab: (1) drop first 4 sheet rows, (2) copy row6 text into blank cells of row5,
     * (3) drop columns whose row6 label is speed/time, (4) drop the row6 line.
     */
    public static PlanInputTabularIo.TabularSheet applyAladdinProcessingPlanDisplaySteps(
            PlanInputTabularIo.TabularSheet raw) {
        Objects.requireNonNull(raw, "raw");
        List<List<String>> rows = new ArrayList<>();
        for (List<String> r : raw.rows()) {
            rows.add(new ArrayList<>(r));
        }
        int maxCol = raw.headers().size();
        for (List<String> r : rows) {
            maxCol = Math.max(maxCol, r.size());
        }
        padRowsToWidth(rows, maxCol);

        int dropHead = Math.min(4, rows.size());
        if (dropHead > 0) {
            rows.subList(0, dropHead).clear();
        }

        if (rows.size() >= 2) {
            List<String> line5 = rows.get(0);
            List<String> line6 = rows.get(1);
            int w = Math.max(line5.size(), line6.size());
            padRowsToWidth(rows, w);
            line5 = rows.get(0);
            line6 = rows.get(1);
            for (int c = 0; c < w; c++) {
                if (isBlankCell(line5.get(c))) {
                    line5.set(c, line6.get(c) != null ? line6.get(c) : "");
                }
            }
            line6 = rows.get(1);
            List<Integer> keep = new ArrayList<>();
            for (int c = 0; c < line6.size(); c++) {
                String label = normalizeAladdinHeaderCell(line6.get(c));
                if (!isAladdinSpeedOrTimeColumn(label)) {
                    keep.add(c);
                }
            }
            if (keep.size() != line6.size()) {
                for (int i = 0; i < rows.size(); i++) {
                    List<String> row = rows.get(i);
                    List<String> next = new ArrayList<>(keep.size());
                    for (int k : keep) {
                        next.add(k < row.size() && row.get(k) != null ? row.get(k) : "");
                    }
                    rows.set(i, next);
                }
            }
            rows.remove(1);
        }

        maxCol = 0;
        for (List<String> r : rows) {
            maxCol = Math.max(maxCol, r.size());
        }
        padRowsToWidth(rows, maxCol);
        List<String> headers = new ArrayList<>();
        for (int c = 0; c < maxCol; c++) {
            headers.add("\u5217" + (c + 1));
        }
        return new PlanInputTabularIo.TabularSheet(headers, rows);
    }

    private static void padRowsToWidth(List<List<String>> rows, int width) {
        for (List<String> r : rows) {
            while (r.size() < width) {
                r.add("");
            }
        }
    }

    private static boolean isBlankCell(String s) {
        return s == null || s.trim().isEmpty();
    }

    private static String normalizeAladdinHeaderCell(String s) {
        return s == null ? "" : s.trim();
    }

    private static boolean isAladdinSpeedOrTimeColumn(String label) {
        return "\u52a0\u5de5\u901f\u5ea6".equals(label) || "\u52a0\u5de5\u6642\u9593".equals(label);
    }
}
