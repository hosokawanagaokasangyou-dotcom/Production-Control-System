package jp.co.pm.ai.kouchin.verify;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.function.Function;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * ①②③原本のコピーシート。要確認キー行を薄黄（FFF2CC）でハイライトする。全行 TableView 化はしない。
 */
public final class SourceRawSheets {

    private static final Set<String> A_KEYS = Set.of(
            Judge.MISMATCH, Judge.NEXT_MONTH, Judge.ONLY_1, Judge.ONLY_2,
            Judge.PREV_ADJUST, Judge.PREV_GAP, Judge.MANUAL_1, Judge.MANUAL_2);
    private static final Set<String> B_KEYS = Set.of(Judge.MISMATCH, Judge.ONLY_2, Judge.ONLY_3);

    private SourceRawSheets() {}

    public static void append(XSSFWorkbook wb, VerifyResult result, Function<Boolean, CellStyle> style) {
        Set<String> attKeiyaku = new HashSet<>();
        for (RecordA rec : result.recordsA()) {
            if (A_KEYS.contains(rec.judge()) && rec.keiyaku() != null && !rec.keiyaku().isBlank()) {
                attKeiyaku.add(Norm.keiyaku(rec.keiyaku()));
            }
        }
        Set<String> attIrai = new HashSet<>();
        for (RecordB rec : result.recordsB()) {
            if (B_KEYS.contains(rec.judge()) && rec.irai() != null && !rec.irai().isBlank()) {
                attIrai.add(Norm.norm(rec.irai()));
            }
        }
        Path p1 = path(result, "①パス");
        Path p2 = path(result, "②パス");
        Path p3 = path(result, "③パス");
        if (p1 != null && Files.isRegularFile(p1)) {
            try {
                List<List<String>> csv = TorayCsvReader.parseCsv(TorayCsvReader.decodeCp932(p1));
                write(wb, "①東レCSV原本", csv, attKeiyaku, attIrai, true, style);
            } catch (RuntimeException ignored) {
            }
        }
        if (p2 != null && Files.isRegularFile(p2)) {
            try (Workbook src = ExcelValues.open(p2)) {
                for (String sheet : result.profile().sheets2()) {
                    List<List<Object>> rows = ExcelValues.readSheet(src, sheet);
                    writeObjects(wb, "②" + sheet + "(当月)", rows, attKeiyaku, attIrai, style);
                }
            } catch (Exception ignored) {
            }
        }
        if (p3 != null && Files.isRegularFile(p3)) {
            try (Workbook src = ExcelValues.open(p3)) {
                writeObjects(wb, "③アラジン原本", ExcelValues.readFirstSheet(src), attKeiyaku, attIrai, style);
            } catch (Exception ignored) {
            }
        }
    }

    private static Path path(VerifyResult r, String key) {
        String s = r.str(key);
        return s.isBlank() ? null : Path.of(s);
    }

    private static void write(
            XSSFWorkbook wb,
            String title,
            List<List<String>> rows,
            Set<String> keiyaku,
            Set<String> irai,
            boolean csv,
            Function<Boolean, CellStyle> style) {
        XSSFSheet ws = wb.createSheet(safeName(wb, title));
        ws.setTabColor(new XSSFColor(new byte[] {(byte) 0xA6, (byte) 0xA6, (byte) 0xA6}, null));
        int r = 0;
        for (List<String> row : rows) {
            boolean hl = hit(row, keiyaku, irai);
            Row xr = ws.createRow(r++);
            for (int c = 0; c < row.size(); c++) {
                Cell cell = xr.createCell(c);
                cell.setCellValue(row.get(c) == null ? "" : row.get(c));
                cell.setCellStyle(style.apply(hl));
            }
        }
        ws.createFreezePane(0, csv ? 1 : 0);
    }

    private static void writeObjects(
            XSSFWorkbook wb,
            String title,
            List<List<Object>> rows,
            Set<String> keiyaku,
            Set<String> irai,
            Function<Boolean, CellStyle> style) {
        XSSFSheet ws = wb.createSheet(safeName(wb, title));
        ws.setTabColor(new XSSFColor(new byte[] {(byte) 0xA6, (byte) 0xA6, (byte) 0xA6}, null));
        int r = 0;
        for (List<Object> row : rows) {
            boolean hl = hitObj(row, keiyaku, irai);
            Row xr = ws.createRow(r++);
            for (int c = 0; c < row.size(); c++) {
                Cell cell = xr.createCell(c);
                Object v = row.get(c);
                if (v instanceof Number n) {
                    cell.setCellValue(n.doubleValue());
                } else {
                    cell.setCellValue(v == null ? "" : String.valueOf(v));
                }
                cell.setCellStyle(style.apply(hl));
            }
        }
        ws.createFreezePane(0, 1);
    }

    private static boolean hit(List<String> row, Set<String> keiyaku, Set<String> irai) {
        int n = Math.min(6, row.size());
        for (int i = 0; i < n; i++) {
            String t = row.get(i);
            if (t == null) {
                continue;
            }
            if (keiyaku.contains(Norm.keiyaku(t)) || irai.contains(Norm.norm(t))) {
                return true;
            }
        }
        return false;
    }

    private static boolean hitObj(List<Object> row, Set<String> keiyaku, Set<String> irai) {
        int n = Math.min(6, row.size());
        for (int i = 0; i < n; i++) {
            String t = Norm.norm(row.get(i));
            if (t.isEmpty()) {
                continue;
            }
            if (keiyaku.contains(Norm.keiyaku(t)) || irai.contains(t)) {
                return true;
            }
        }
        return false;
    }

    private static String safeName(Workbook wb, String title) {
        String base = title.length() > 31 ? title.substring(0, 31) : title;
        String name = base;
        int i = 1;
        while (wb.getSheet(name) != null) {
            String suffix = "_" + i;
            name = (base.length() + suffix.length() > 31 ? base.substring(0, 31 - suffix.length()) : base) + suffix;
            i++;
        }
        return name;
    }
}
