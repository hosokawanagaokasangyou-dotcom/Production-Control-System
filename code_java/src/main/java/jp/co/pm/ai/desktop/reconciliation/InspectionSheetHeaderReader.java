package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.Locale;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import jp.co.pm.ai.desktop.io.PoiWorkbookOpener;

/** 検査表ブック上側から依頼NOと加工日を読む。 */
public final class InspectionSheetHeaderReader {

    private static final int MAX_ROWS = 10;
    private static final int MAX_COLS = 50;
    private static final DataFormatter FORMATTER = new DataFormatter(Locale.JAPAN);

    public record Header(String iraiNo, LocalDate processingDate) {}

    private InspectionSheetHeaderReader() {}

    public static Header read(Path file) throws IOException {
        if (file == null) {
            return new Header("", null);
        }
        try (Workbook wb = PoiWorkbookOpener.open(file.toFile())) {
            Sheet sheet = selectSheet(wb);
            if (sheet == null) {
                return fallbackFromFileName(file);
            }
            String irai = "";
            LocalDate date = null;
            for (int r = 0; r < MAX_ROWS; r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    continue;
                }
                for (int c = 0; c < MAX_COLS; c++) {
                    Cell cell = row.getCell(c);
                    String text = cellText(cell);
                    if (irai.isEmpty() && isIraiLabel(text)) {
                        irai = firstIraiToTheRight(row, c);
                    }
                    if (date == null && InspectionSheetProcessingDate.isProcessingDateLabel(text)) {
                        Optional<LocalDate> parsed = InspectionSheetProcessingDate.parse(text);
                        if (parsed.isPresent()) {
                            date = parsed.get();
                        } else {
                            date = dateToTheRight(row, c).orElse(null);
                        }
                    }
                }
            }
            if (irai.isEmpty()) {
                irai = InspectionSheetIraiNo.extractFromFileName(fileName(file)).orElse("");
            }
            return new Header(irai, date);
        } catch (IOException ex) {
            throw ex;
        } catch (RuntimeException ex) {
            Header fallback = fallbackFromFileName(file);
            if (!fallback.iraiNo().isEmpty()) {
                return fallback;
            }
            throw new IOException("検査表ヘッダ読取に失敗: " + file.getFileName() + ": " + ex.getMessage(), ex);
        }
    }

    private static Header fallbackFromFileName(Path file) {
        return new Header(InspectionSheetIraiNo.extractFromFileName(fileName(file)).orElse(""), null);
    }

    private static String fileName(Path file) {
        Path name = file.getFileName();
        return name != null ? name.toString() : "";
    }

    static Sheet selectSheet(Workbook wb) {
        if (wb == null || wb.getNumberOfSheets() <= 0) {
            return null;
        }
        Sheet named = wb.getSheet("検査表");
        if (named != null) {
            return named;
        }
        return wb.getSheetAt(0);
    }

    static boolean isIraiLabel(String text) {
        if (text == null || text.isBlank()) {
            return false;
        }
        String half = InspectionSheetIraiNo.toHalfWidth(text);
        if (half.contains("加工依頼")) {
            return true;
        }
        String lower = half.toLowerCase(Locale.ROOT);
        return lower.contains("依頼") && (lower.contains("no") || half.contains("№"));
    }

    private static String firstIraiToTheRight(Row row, int labelCol) {
        for (int c = labelCol + 1; c < MAX_COLS; c++) {
            Optional<String> irai = InspectionSheetIraiNo.extractFromCellText(cellText(row.getCell(c)));
            if (irai.isPresent()) {
                return irai.get();
            }
        }
        return "";
    }

    private static Optional<LocalDate> dateToTheRight(Row row, int labelCol) {
        for (int c = labelCol + 1; c < MAX_COLS; c++) {
            Cell cell = row.getCell(c);
            if (cell == null) {
                continue;
            }
            switch (cell.getCellType()) {
                case NUMERIC -> {
                    Optional<LocalDate> serial =
                            InspectionSheetProcessingDate.fromExcelSerial(cell.getNumericCellValue());
                    if (serial.isPresent()) {
                        return serial;
                    }
                }
                default -> {
                    Optional<LocalDate> parsed = InspectionSheetProcessingDate.parse(cellText(cell));
                    if (parsed.isPresent()) {
                        return parsed;
                    }
                }
            }
        }
        return Optional.empty();
    }

    private static String cellText(Cell cell) {
        if (cell == null) {
            return "";
        }
        return FORMATTER.formatCellValue(cell);
    }
}
