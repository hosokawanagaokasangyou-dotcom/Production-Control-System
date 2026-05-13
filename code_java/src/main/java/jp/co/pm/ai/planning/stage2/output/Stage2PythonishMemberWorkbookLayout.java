package jp.co.pm.ai.planning.stage2.output;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Optional;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Python 段階2の人員ブック（メンバー名＝シート名、1 列目「時間帯」＋暦日列）に近いプレースホルダ。
 * 暦日は実行日 1 列のみ（未配台のため中身は空）。
 */
public final class Stage2PythonishMemberWorkbookLayout {

    private static final DateTimeFormatter DAY_HEADER =
            DateTimeFormatter.ofPattern("MM/dd (EEE)", Locale.US);

    private Stage2PythonishMemberWorkbookLayout() {}

    public static void write(
            Path path,
            List<String> memberDisplayNames,
            LocalDate anchorDay,
            Optional<LocalTime> factoryStart,
            Optional<LocalTime> factoryEnd)
            throws IOException {
        Files.createDirectories(path.getParent());
        LocalTime start = factoryStart.orElse(LocalTime.of(8, 45));
        LocalTime end = factoryEnd.orElse(LocalTime.of(17, 0));
        if (!end.isAfter(start)) {
            end = start.plusHours(1);
        }
        String dayCol = anchorDay.format(DAY_HEADER);
        List<String> headers = new ArrayList<>();
        headers.add("時間帯");
        headers.add(dayCol);
        List<List<String>> slotRows = buildSlotRows(start, end);

        List<String> members =
                memberDisplayNames != null && !memberDisplayNames.isEmpty()
                        ? memberDisplayNames
                        : List.of("member");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Set<String> used = new HashSet<>();
            for (String m : members) {
                String base = sanitizeSheetName(m);
                String name = base;
                int i = 2;
                while (used.contains(name)) {
                    name = base + "_" + i++;
                }
                used.add(name);
                Sheet sh = wb.createSheet(name);
                writeTabular(sh, headers, slotRows);
            }
            try (OutputStream os = Files.newOutputStream(path)) {
                wb.write(os);
            }
        }
    }

    private static List<List<String>> buildSlotRows(LocalTime start, LocalTime end) {
        List<List<String>> rows = new ArrayList<>();
        LocalTime t = start;
        while (t.isBefore(end)) {
            LocalTime next = t.plusMinutes(10);
            if (next.isAfter(end)) {
                next = end;
            }
            rows.add(List.of(slotLabel(t, next), ""));
            t = next;
            if (rows.size() > 5000) {
                break;
            }
        }
        return rows;
    }

    private static String slotLabel(LocalTime t0, LocalTime t1) {
        return String.format(
                Locale.ROOT, "%02d:%02d-%02d:%02d", t0.getHour(), t0.getMinute(), t1.getHour(), t1.getMinute());
    }

    private static String sanitizeSheetName(String name) {
        String n = name == null ? "member" : name.strip();
        if (n.isEmpty()) {
            n = "member";
        }
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < n.length() && sb.length() < 28; i++) {
            char c = n.charAt(i);
            if (c == '[' || c == ']' || c == '*' || c == '/' || c == '\\' || c == '?') {
                sb.append('_');
            } else {
                sb.append(c);
            }
        }
        return sb.toString();
    }

    private static void writeTabular(Sheet sh, List<String> headers, List<List<String>> rows) {
        Row hr = sh.createRow(0);
        for (int c = 0; c < headers.size(); c++) {
            Cell cell = hr.createCell(c);
            String v = headers.get(c);
            cell.setCellValue(v != null ? v : "");
        }
        int r = 1;
        for (List<String> rowVals : rows) {
            Row rr = sh.createRow(r++);
            for (int c = 0; c < headers.size(); c++) {
                Cell cell = rr.createCell(c);
                String v = c < rowVals.size() && rowVals.get(c) != null ? rowVals.get(c) : "";
                cell.setCellValue(v);
            }
        }
    }
}
