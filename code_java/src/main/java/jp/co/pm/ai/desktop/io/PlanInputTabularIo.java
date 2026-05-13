package jp.co.pm.ai.desktop.io;

import java.io.BufferedReader;
import java.io.BufferedWriter;
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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Read/write the stage-2 plan-input sheet as tabular data (parity with {@code load_planning_tasks_df} /
 * {@code PM_AI_PLAN_INPUT_PATH}). Supports {@code .csv} (UTF-8) and Excel {@code .xlsx/.xlsm}.
 */
public final class PlanInputTabularIo {

    private PlanInputTabularIo() {}

    public record TabularSheet(List<String> headers, List<List<String>> rows) {}

    /**
     * 読み取った表と、実際に開いた Excel シート名（CSV のときは空文字）。段階1出力の {@code plan_input_tasks.xlsx}
     * はシート名が「タスク一覧」のみのことがあり、Python の {@code _resolve_tabular_sheet_name_calamine} と同様に
     * 要求シートが無くシートが1枚だけのときはそのシートを採用する。
     */
    public record TabularRead(String resolvedSheetName, TabularSheet tabular) {}

    public static TabularSheet read(Path path, String sheetName) throws IOException {
        return readWithResolvedSheet(path, sheetName).tabular();
    }

    public static TabularRead readWithResolvedSheet(Path path, String sheetName) throws IOException {
        String low = path.getFileName().toString().toLowerCase(Locale.ROOT);
        if (low.endsWith(".csv")) {
            return new TabularRead("", readCsv(path));
        }
        if (low.endsWith(".xlsx") || low.endsWith(".xlsm")) {
            return readExcelResolved(path, sheetName);
        }
        throw new IOException("unsupported extension (use .csv, .xlsx, .xlsm): " + path);
    }

    public static void write(Path path, String sheetName, TabularSheet data) throws IOException {
        if (data == null || data.headers() == null) {
            throw new IOException("no data");
        }
        String low = path.getFileName().toString().toLowerCase(Locale.ROOT);
        if (low.endsWith(".csv")) {
            writeCsv(path, data);
            return;
        }
        if (low.endsWith(".xlsx") || low.endsWith(".xlsm")) {
            writeExcel(path, sheetName, data);
            return;
        }
        throw new IOException("unsupported extension (use .csv, .xlsx, .xlsm): " + path);
    }

    private static TabularSheet readCsv(Path path) throws IOException {
        List<String> headers = new ArrayList<>();
        List<List<String>> rows = new ArrayList<>();
        try (BufferedReader r = Files.newBufferedReader(path, StandardCharsets.UTF_8)) {
            String headerLine = r.readLine();
            if (headerLine == null) {
                return new TabularSheet(headers, rows);
            }
            parseCsvLine(headerLine, headers);
            if (!headers.isEmpty()) {
                String h0 = headers.get(0);
                if (!h0.isEmpty() && h0.charAt(0) == '﻿') {
                    headers.set(0, h0.substring(1));
                }
            }
            String line;
            while ((line = r.readLine()) != null) {
                List<String> cells = new ArrayList<>();
                parseCsvLine(line, cells);
                while (cells.size() < headers.size()) {
                    cells.add("");
                }
                if (cells.size() > headers.size()) {
                    cells.subList(headers.size(), cells.size()).clear();
                }
                rows.add(cells);
            }
        }
        return new TabularSheet(headers, rows);
    }

    /** Minimal CSV: comma-separated, double-quote escape. */
    static void parseCsvLine(String line, List<String> out) {
        out.clear();
        if (line.isEmpty()) {
            return;
        }
        StringBuilder cur = new StringBuilder();
        boolean inQ = false;
        for (int i = 0; i < line.length(); i++) {
            char c = line.charAt(i);
            if (inQ) {
                if (c == '"') {
                    if (i + 1 < line.length() && line.charAt(i + 1) == '"') {
                        cur.append('"');
                        i++;
                    } else {
                        inQ = false;
                    }
                } else {
                    cur.append(c);
                }
            } else {
                if (c == '"') {
                    inQ = true;
                } else if (c == ',') {
                    out.add(cur.toString());
                    cur.setLength(0);
                } else {
                    cur.append(c);
                }
            }
        }
        out.add(cur.toString());
    }

    private static void writeCsv(Path path, TabularSheet data) throws IOException {
        try (BufferedWriter w = Files.newBufferedWriter(path, StandardCharsets.UTF_8)) {
            w.write('﻿');
            w.write(joinCsvLine(data.headers()));
            w.newLine();
            for (List<String> row : data.rows()) {
                List<String> cells = new ArrayList<>(row);
                while (cells.size() < data.headers().size()) {
                    cells.add("");
                }
                if (cells.size() > data.headers().size()) {
                    cells = cells.subList(0, data.headers().size());
                }
                w.write(joinCsvLine(cells));
                w.newLine();
            }
        }
    }

    private static String joinCsvLine(List<String> cells) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < cells.size(); i++) {
            if (i > 0) {
                sb.append(',');
            }
            String s = cells.get(i) != null ? cells.get(i) : "";
            if (s.contains(",") || s.contains("\"") || s.contains("\n") || s.contains("\r")) {
                sb.append('"').append(s.replace("\"", "\"\"")).append('"');
            } else {
                sb.append(s);
            }
        }
        return sb.toString();
    }

    private static TabularRead readExcelResolved(Path path, String sheetName) throws IOException {
        try (Workbook wb = WorkbookFactory.create(path.toFile())) {
            Sheet sh = wb.getSheet(sheetName);
            String usedSheetName = sheetName;
            if (sh == null && wb.getNumberOfSheets() == 1) {
                sh = wb.getSheetAt(0);
                usedSheetName = sh.getSheetName();
            }
            if (sh == null) {
                List<String> names = new ArrayList<>();
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    names.add(wb.getSheetName(i));
                }
                throw new IOException(
                        "sheet not found: \""
                                + sheetName
                                + "\" in "
                                + path
                                + "; sheets="
                                + names);
            }
            Row h = sh.getRow(0);
            if (h == null) {
                return new TabularRead(usedSheetName, new TabularSheet(List.of(), List.of()));
            }
            List<String> headers = new ArrayList<>();
            short last = h.getLastCellNum();
            for (int c = 0; c < last; c++) {
                headers.add(cellToString(h.getCell(c)));
            }
            trimTrailingBlankHeaders(headers);
            List<List<String>> rows = new ArrayList<>();
            for (int r = 1; r <= sh.getLastRowNum(); r++) {
                Row row = sh.getRow(r);
                List<String> line = new ArrayList<>(headers.size());
                for (int c = 0; c < headers.size(); c++) {
                    line.add(row == null ? "" : cellToString(row.getCell(c)));
                }
                rows.add(line);
            }
            return new TabularRead(usedSheetName, new TabularSheet(headers, rows));
        }
    }

    private static void writeExcel(Path path, String sheetName, TabularSheet data) throws IOException {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sh = wb.createSheet(sheetName);
            Row hr = sh.createRow(0);
            for (int c = 0; c < data.headers().size(); c++) {
                Cell cell = hr.createCell(c);
                String v = data.headers().get(c);
                cell.setCellValue(v != null ? v : "");
            }
            int r = 1;
            for (List<String> rowVals : data.rows()) {
                Row rr = sh.createRow(r++);
                for (int c = 0; c < data.headers().size(); c++) {
                    Cell cell = rr.createCell(c);
                    String v = c < rowVals.size() && rowVals.get(c) != null ? rowVals.get(c) : "";
                    cell.setCellValue(v);
                }
            }
            try (var out = Files.newOutputStream(path)) {
                wb.write(out);
            }
        }
    }

    private static void trimTrailingBlankHeaders(List<String> headers) {
        while (!headers.isEmpty() && headers.get(headers.size() - 1).isBlank()) {
            headers.remove(headers.size() - 1);
        }
    }

    private static String cellToString(Cell cell) {
        return ExcelCellReadSupport.cellToDisplayString(cell);
    }
}
