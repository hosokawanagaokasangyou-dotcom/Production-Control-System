package jp.co.pm.ai.planning.stage2.input;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import jp.co.pm.ai.desktop.io.ExcelCellReadSupport;

/**
 * master の「メイン」シート A12/B12（稼働開始・終了）を読む。Python {@code _read_master_main_factory_operating_times}
 * に概ね整合。
 */
public final class Stage2MasterFactoryHoursReader {

    private Stage2MasterFactoryHoursReader() {}

    public record FactoryHours(Optional<LocalTime> start, Optional<LocalTime> end) {}

    public static FactoryHours read(Path masterPath) throws IOException {
        if (masterPath == null || !Files.isRegularFile(masterPath)) {
            return new FactoryHours(Optional.empty(), Optional.empty());
        }
        DataFormatter fmt = new DataFormatter(Locale.ROOT);
        try (Workbook wb = WorkbookFactory.create(masterPath.toFile())) {
            String sheetName = pickMainSheetName(wb);
            if (sheetName == null) {
                return new FactoryHours(Optional.empty(), Optional.empty());
            }
            Sheet sh = wb.getSheet(sheetName);
            if (sh == null) {
                return new FactoryHours(Optional.empty(), Optional.empty());
            }
            Row r = sh.getRow(11);
            if (r == null) {
                return new FactoryHours(Optional.empty(), Optional.empty());
            }
            Optional<LocalTime> st = cellToTimeOptional(r.getCell(0), fmt);
            Optional<LocalTime> et = cellToTimeOptional(r.getCell(1), fmt);
            if (st.isEmpty() || et.isEmpty()) {
                return new FactoryHours(Optional.empty(), Optional.empty());
            }
            if (!st.get().isBefore(et.get())) {
                return new FactoryHours(Optional.empty(), Optional.empty());
            }
            return new FactoryHours(st, et);
        }
    }

    private static String pickMainSheetName(Workbook wb) {
        List<String> cand = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String sn = wb.getSheetName(i);
            if (sn != null && sn.contains("メイン") && !sn.contains("カレンダー")) {
                cand.add(sn);
            }
        }
        if (cand.isEmpty()) {
            return null;
        }
        return cand.stream().min(Comparator.comparingInt(String::length)).orElse(null);
    }

    private static Optional<LocalTime> cellToTimeOptional(Cell cell, DataFormatter fmt) {
        if (cell == null) {
            return Optional.empty();
        }
        try {
            if (DateUtil.isCellDateFormatted(cell)) {
                Date d = cell.getDateCellValue();
                if (d == null) {
                    return Optional.empty();
                }
                return Optional.of(
                        d.toInstant().atZone(ZoneId.systemDefault()).toLocalTime());
            }
        } catch (Exception ignored) {
            // fall through to string
        }
        String s = ExcelCellReadSupport.normalizeCommaDigitArtifacts(fmt.formatCellValue(cell).strip());
        if (s.isEmpty() || "nan".equalsIgnoreCase(s)) {
            return Optional.empty();
        }
        for (String sep : List.of(":", "：")) {
            int idx = s.indexOf(sep.charAt(0));
            if (idx > 0) {
                try {
                    int h = Integer.parseInt(s.substring(0, idx).strip());
                    int m = Integer.parseInt(s.substring(idx + 1).strip());
                    return Optional.of(LocalTime.of(h, m));
                } catch (NumberFormatException ignored) {
                    return Optional.empty();
                }
            }
        }
        return Optional.empty();
    }
}
