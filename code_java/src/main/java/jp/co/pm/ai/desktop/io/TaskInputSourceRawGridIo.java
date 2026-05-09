package jp.co.pm.ai.desktop.io;

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Raw sheet reader for {@link jp.co.pm.ai.desktop.config.AppPaths#KEY_PM_AI_TASK_INPUT_SOURCE_DIR}.
 * All sheet rows are data; synthetic headers use column index labels (see {@code readRaw} output).
 */
public final class TaskInputSourceRawGridIo {

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
                line.add(
                        ExcelCellReadSupport.normalizeCommaDigitArtifacts(
                                c < src.size() && src.get(c) != null ? src.get(c) : ""));
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
        return ExcelCellReadSupport.cellToDisplayString(cell);
    }

    /**
     * Aladdin tab: (1) drop first 4 sheet rows, (2) copy row6 text into blank cells of row5,
     * (3) drop columns whose row6 label is speed/time, (4) drop the row6 line,
     * (5) in the top row, normalize date-like cells to {@code yyyy/MM/dd} using year from the
     * first data row's \u53d7\u6ce8\u65e5 column, (6) use that top row as headers and remove it from data.
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
        if (!rows.isEmpty()) {
            normalizeAladdinHeaderRowDateCells(rows);
        }
        if (rows.isEmpty()) {
            return new PlanInputTabularIo.TabularSheet(List.of(), rows);
        }
        List<String> headerRow = rows.get(0);
        List<String> headers = new ArrayList<>(headerRow.size());
        for (String cell : headerRow) {
            headers.add(cell != null ? cell : "");
        }
        rows.remove(0);
        return new PlanInputTabularIo.TabularSheet(headers, rows);
    }

    /**
     * Processing actuals tab: (1) drop the first 4 sheet rows, (2) treat the next row (original 5th row,
     * 1-based) as column headers and remove it from the data body.
     */
    public static PlanInputTabularIo.TabularSheet applyProcessingActualsDisplaySteps(
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

        if (rows.isEmpty()) {
            return new PlanInputTabularIo.TabularSheet(List.of(), List.of());
        }

        List<String> headerRow = rows.get(0);
        List<String> headers = new ArrayList<>(headerRow.size());
        for (String cell : headerRow) {
            headers.add(cell != null ? cell : "");
        }
        rows.remove(0);
        return new PlanInputTabularIo.TabularSheet(headers, rows);
    }

    private static final DateTimeFormatter PROCESSING_DATETIME_OUT =
            DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm");

    /** Workbook header for process name (dedupe key). */
    private static final String HEADER_PROCESS_NAME = "\u5de5\u7a0b\u540d";

    private static final String HEADER_MACHINE_NAME = "\u6a5f\u68b0\u540d";

    /** Request-id column title (halfwidth Latin suffix). */
    private static final String HEADER_REQUEST_NO_ASCII = "\u4f9d\u983cNO";

    /** Alternate request-id header (fullwidth Latin) when {@link #HEADER_REQUEST_NO_ASCII} is absent. */
    private static final String HEADER_REQUEST_NO_FULL = "\u4f9d\u983c\uff2e\uff2f";

    private static final String HEADER_KAKOU_DATE = "\u52a0\u5de5\u65e5";
    private static final String HEADER_START_HOUR = "\u958b\u59cb\u6642\u9593";
    private static final String HEADER_START_MIN = "\u958b\u59cb\u5206";
    private static final String HEADER_END_HOUR = "\u7d42\u4e86\u6642\u9593";
    private static final String HEADER_END_MIN = "\u7d42\u4e86\u5206";
    private static final String HEADER_KAKOU_START_DT = "\u52a0\u5de5\u958b\u59cb\u65e5\u6642";
    private static final String HEADER_KAKOU_END_DT = "\u52a0\u5de5\u7d42\u4e86\u65e5\u6642";

    /**
     * Appends {@code ??????} and {@code ??????} from {@code ???} + ({@code ????},{@code ???})
     * and {@code ???} + ({@code ????},{@code ???}). Output uses {@link #PROCESSING_DATETIME_OUT}.
     * Missing source headers or unparseable cells yield empty appended cells.
     */
    public static PlanInputTabularIo.TabularSheet applyProcessingActualsDateTimeColumns(
            PlanInputTabularIo.TabularSheet shaped) {
        Objects.requireNonNull(shaped, "shaped");
        List<String> headers = new ArrayList<>(shaped.headers());
        int idxDate = indexOfHeaderTitle(headers, HEADER_KAKOU_DATE);
        int idxSh = indexOfHeaderTitle(headers, HEADER_START_HOUR);
        int idxSm = indexOfHeaderTitle(headers, HEADER_START_MIN);
        int idxEh = indexOfHeaderTitle(headers, HEADER_END_HOUR);
        int idxEm = indexOfHeaderTitle(headers, HEADER_END_MIN);

        int baseCols = headers.size();
        List<List<String>> outRows = new ArrayList<>();
        for (List<String> src : shaped.rows()) {
            List<String> row = new ArrayList<>(src);
            while (row.size() < baseCols) {
                row.add("");
            }
            String startDt = formatProcessingDateTimeCell(row, idxDate, idxSh, idxSm);
            String endDt = formatProcessingDateTimeCell(row, idxDate, idxEh, idxEm);
            row.add(startDt);
            row.add(endDt);
            outRows.add(row);
        }
        headers.add(HEADER_KAKOU_START_DT);
        headers.add(HEADER_KAKOU_END_DT);
        return new PlanInputTabularIo.TabularSheet(headers, outRows);
    }

    /**
     * Keeps the first row per key (process name, machine name, request no., processing date); drops later
     * duplicates. If any of those header columns is absent, returns the input unchanged.
     */
    public static PlanInputTabularIo.TabularSheet applyProcessingActualsDedupeByQuadKey(
            PlanInputTabularIo.TabularSheet shaped) {
        Objects.requireNonNull(shaped, "shaped");
        List<String> headers = shaped.headers();
        int iProc = indexOfHeaderTitle(headers, HEADER_PROCESS_NAME);
        int iMach = indexOfHeaderTitle(headers, HEADER_MACHINE_NAME);
        int iReq = indexOfHeaderFirst(headers, HEADER_REQUEST_NO_ASCII, HEADER_REQUEST_NO_FULL);
        int iDate = indexOfHeaderTitle(headers, HEADER_KAKOU_DATE);
        if (iProc < 0 || iMach < 0 || iReq < 0 || iDate < 0) {
            return shaped;
        }
        Set<String> seen = new HashSet<>();
        List<List<String>> outRows = new ArrayList<>();
        for (List<String> src : shaped.rows()) {
            String key = quadDedupeKey(src, iProc, iMach, iReq, iDate);
            if (seen.add(key)) {
                outRows.add(new ArrayList<>(src));
            }
        }
        return new PlanInputTabularIo.TabularSheet(new ArrayList<>(headers), outRows);
    }

    private static int indexOfHeaderFirst(List<String> headers, String... titles) {
        if (titles == null) {
            return -1;
        }
        for (String t : titles) {
            int ix = indexOfHeaderTitle(headers, t);
            if (ix >= 0) {
                return ix;
            }
        }
        return -1;
    }

    private static String quadDedupeKey(
            List<String> row, int iProc, int iMach, int iReq, int iDate) {
        return cellAt(row, iProc).strip()
                + '\u001e'
                + cellAt(row, iMach).strip()
                + '\u001e'
                + cellAt(row, iReq).strip()
                + '\u001e'
                + cellAt(row, iDate).strip();
    }

    private static int indexOfHeaderTitle(List<String> headers, String title) {
        if (headers == null || title == null) {
            return -1;
        }
        for (int i = 0; i < headers.size(); i++) {
            String h = headers.get(i);
            if (title.equals(h != null ? h.strip() : "")) {
                return i;
            }
        }
        return -1;
    }

    private static String formatProcessingDateTimeCell(
            List<String> row, int idxDate, int idxHour, int idxMin) {
        if (idxDate < 0 || idxHour < 0 || idxMin < 0) {
            return "";
        }
        LocalDate d = parseKakouDate(cellAt(row, idxDate));
        int hh = parseHourMinutePart(cellAt(row, idxHour));
        int mm = parseHourMinutePart(cellAt(row, idxMin));
        if (d == null || hh < 0 || mm < 0 || hh > 23 || mm > 59) {
            return "";
        }
        try {
            return LocalDateTime.of(d, LocalTime.of(hh, mm)).format(PROCESSING_DATETIME_OUT);
        } catch (Exception e) {
            return "";
        }
    }

    private static String cellAt(List<String> row, int idx) {
        if (idx < 0 || idx >= row.size()) {
            return "";
        }
        String s = row.get(idx);
        return s != null ? s : "";
    }

    private static LocalDate parseKakouDate(String raw) {
        if (isBlankCell(raw)) {
            return null;
        }
        String trimmed = raw.trim();
        String withoutWeekday = stripTrailingWeekdayInParens(trimmed);
        LocalDate d = tryParseFlexibleDate(withoutWeekday);
        if (d != null) {
            return d;
        }
        return tryParseFlexibleDate(trimmed);
    }

    private static int parseHourMinutePart(String raw) {
        if (raw == null) {
            return -1;
        }
        String t = raw.strip().replace(",", "");
        if (t.isEmpty()) {
            return -1;
        }
        int dot = t.indexOf('.');
        if (dot >= 0) {
            t = t.substring(0, dot);
        }
        try {
            return Integer.parseInt(t.strip());
        } catch (NumberFormatException e) {
            return -1;
        }
    }

    private static final DateTimeFormatter ALADDIN_HEADER_DATE_OUT =
            DateTimeFormatter.ofPattern("yyyy/MM/dd");

    private static final List<DateTimeFormatter> FLEX_DATE_IN =
            List.of(
                    DateTimeFormatter.ofPattern("yyyy/M/d"),
                    DateTimeFormatter.ofPattern("yyyy/MM/dd"),
                    DateTimeFormatter.ofPattern("yyyy/M/dd"),
                    DateTimeFormatter.ofPattern("yyyy/MM/d"),
                    DateTimeFormatter.ofPattern("yyyy.M.d"),
                    DateTimeFormatter.ofPattern("yyyy.MM.dd"));

    /** Month/day only (no year), aligned with first data row year from \u53d7\u6ce8\u65e5 column. */
    private static final Pattern MONTH_DAY_SLASH = Pattern.compile("^\\s*(\\d{1,2})/(\\d{1,2})\\s*$");

    /**
     * Top row = future headers: cells that look like dates become {@code yyyy/MM/dd} (e.g. {@code 2026/04/01}).
     * Trailing weekday suffixes like {@code (?)} are stripped before parsing. Year for month/day-only
     * values is taken from the first data row at the column whose header cell is \u53d7\u6ce8\u65e5.
     * The literal header text cell \u53d7\u6ce8\u65e5 is left unchanged.
     */
    private static void normalizeAladdinHeaderRowDateCells(List<List<String>> rows) {
        if (rows.size() < 2) {
            return;
        }
        int mc = 0;
        for (List<String> r : rows) {
            mc = Math.max(mc, r.size());
        }
        padRowsToWidth(rows, mc);
        List<String> top = rows.get(0);
        int jIdx = indexOfJuchuBiColumn(top);
        if (jIdx < 0) {
            return;
        }
        List<String> firstData = rows.get(1);
        String jVal = jIdx < firstData.size() && firstData.get(jIdx) != null ? firstData.get(jIdx) : "";
        int year = extractYearFromJuchuDataCell(jVal);
        if (year < 0) {
            return;
        }
        for (int c = 0; c < top.size(); c++) {
            if (c == jIdx) {
                continue;
            }
            String raw = top.get(c) != null ? top.get(c) : "";
            String out = formatAladdinHeaderDateCell(raw, year);
            if (out != null) {
                top.set(c, out);
            }
        }
    }

    private static int indexOfJuchuBiColumn(List<String> topRow) {
        for (int i = 0; i < topRow.size(); i++) {
            if ("\u53d7\u6ce8\u65e5".equals(normalizeAladdinHeaderCell(topRow.get(i)))) {
                return i;
            }
        }
        return -1;
    }

    private static int extractYearFromJuchuDataCell(String s) {
        if (isBlankCell(s)) {
            return -1;
        }
        LocalDate d = tryParseFlexibleDate(s.trim());
        if (d != null) {
            return d.getYear();
        }
        Matcher m = Pattern.compile("(20[0-9]{2})").matcher(s.trim());
        if (m.find()) {
            try {
                return Integer.parseInt(m.group(1));
            } catch (NumberFormatException ignored) {
                return -1;
            }
        }
        return -1;
    }

    private static LocalDate tryParseFlexibleDate(String t) {
        for (DateTimeFormatter f : FLEX_DATE_IN) {
            try {
                return LocalDate.parse(t, f);
            } catch (DateTimeParseException ignored) {
                // next
            }
        }
        return null;
    }

    /** Trailing {@code (?)} / {@code (?)} labels on calendar-style headers (e.g. {@code 05/25(?)}). */
    private static String stripTrailingWeekdayInParens(String s) {
        if (s == null || s.isEmpty()) {
            return "";
        }
        return s.replaceFirst("\\s*\\([^)]*\\)\\s*$", "").trim();
    }

    /** Returns formatted date or null if left unchanged. Output always uses {@link #ALADDIN_HEADER_DATE_OUT}. */
    private static String formatAladdinHeaderDateCell(String raw, int year) {
        if (isBlankCell(raw)) {
            return null;
        }
        String t = raw.trim();
        if ("\u53d7\u6ce8\u65e5".equals(t)) {
            return null;
        }
        String core = stripTrailingWeekdayInParens(t);
        LocalDate full = tryParseFlexibleDate(t);
        if (full == null) {
            full = tryParseFlexibleDate(core);
        }
        if (full != null) {
            return full.format(ALADDIN_HEADER_DATE_OUT);
        }
        Matcher md = MONTH_DAY_SLASH.matcher(core);
        if (md.matches()) {
            int month = Integer.parseInt(md.group(1));
            int day = Integer.parseInt(md.group(2));
            try {
                return LocalDate.of(year, month, day).format(ALADDIN_HEADER_DATE_OUT);
            } catch (Exception ignored) {
                return null;
            }
        }
        return null;
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
