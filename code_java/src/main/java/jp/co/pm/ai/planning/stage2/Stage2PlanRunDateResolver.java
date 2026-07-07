package jp.co.pm.ai.planning.stage2;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.dispatch.AladdinShapedPlanQtyLookup;
import jp.co.pm.ai.desktop.io.ExcelCellReadSupport;

/**
 * 段階2の「当日」暦日（Python {@code _extract_data_extraction_datetime} の date 部分、skip_today 適用前）。
 */
public final class Stage2PlanRunDateResolver {

    private static final DateTimeFormatter[] DATE_TIME_FORMATS =
            new DateTimeFormatter[] {
                DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss"),
                DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"),
                DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm"),
                DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm"),
            };

    private Stage2PlanRunDateResolver() {}

    public static LocalDate resolvePlanDate(Map<String, String> ui) {
        Optional<LocalDateTime> dt = extractDataExtractionDateTime(ui);
        return dt.map(LocalDateTime::toLocalDate).orElseGet(LocalDate::now);
    }

    /** Aladdin shaped JSON の日付列キー（{@code yyyy/MM/dd}）。 */
    public static String planDateColumnKey(Map<String, String> ui) {
        LocalDate d = resolvePlanDate(ui);
        return d.format(DateTimeFormatter.ofPattern("yyyy/MM/dd"));
    }

    static Optional<LocalDateTime> extractDataExtractionDateTime(Map<String, String> ui) {
        Path workbook = resolveDataExtractionWorkbook(ui);
        if (workbook == null || !Files.isRegularFile(workbook)) {
            return Optional.empty();
        }
        String sheetName = "加工計画DATA";
        try (Workbook wb = WorkbookFactory.create(workbook.toFile())) {
            Sheet sheet = wb.getSheet(sheetName);
            if (sheet == null) {
                return Optional.empty();
            }
            for (String colName :
                    List.of("データ抽出時間", "抽出時間", "データ抽出日")) {
                Optional<LocalDateTime> parsed = firstDateTimeInColumn(sheet, colName);
                if (parsed.isPresent()) {
                    return parsed;
                }
            }
        } catch (IOException ignored) {
            return Optional.empty();
        }
        return Optional.empty();
    }

    private static Path resolveDataExtractionWorkbook(Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String fromEnv = u.get(AppPaths.KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK);
        if (fromEnv != null && !fromEnv.isBlank()) {
            Path p = Path.of(fromEnv.strip()).toAbsolutePath().normalize();
            if (Files.isRegularFile(p)) {
                return p;
            }
        }
        String planInput = u.get(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH);
        if (planInput != null && !planInput.isBlank()) {
            Path p = Path.of(planInput.strip()).toAbsolutePath().normalize();
            if (Files.isRegularFile(p)) {
                return p;
            }
        }
        return null;
    }

    private static Optional<LocalDateTime> firstDateTimeInColumn(Sheet sheet, String headerName) {
        Row header = sheet.getRow(sheet.getFirstRowNum());
        if (header == null) {
            return Optional.empty();
        }
        int colIdx = -1;
        short last = header.getLastCellNum();
        for (int c = 0; c < last; c++) {
            String h = ExcelCellReadSupport.cellToDisplayString(header.getCell(c));
            if (headerName.equals(h != null ? h.strip() : "")) {
                colIdx = c;
                break;
            }
        }
        if (colIdx < 0) {
            return Optional.empty();
        }
        int lastRow = sheet.getLastRowNum();
        for (int r = sheet.getFirstRowNum() + 1; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            String text =
                    ExcelCellReadSupport.stripMidnightDateTimeSuffix(
                            ExcelCellReadSupport.cellToDisplayString(row.getCell(colIdx)));
            if (text == null || text.isBlank()) {
                continue;
            }
            Optional<LocalDateTime> parsed = parseDateTime(text.strip());
            if (parsed.isPresent()) {
                return parsed;
            }
        }
        return Optional.empty();
    }

    private static Optional<LocalDateTime> parseDateTime(String text) {
        for (DateTimeFormatter fmt : DATE_TIME_FORMATS) {
            try {
                return Optional.of(LocalDateTime.parse(text, fmt));
            } catch (DateTimeParseException ignored) {
                // next
            }
        }
        LocalDate d = AladdinShapedPlanQtyLookup.parsePlanDateColumn(text);
        if (d != null) {
            return Optional.of(d.atStartOfDay());
        }
        return Optional.empty();
    }
}
