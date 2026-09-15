package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.Locale;
import java.util.Optional;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import jp.co.pm.ai.desktop.io.PoiWorkbookOpener;

/** 検査表ブック上側から依頼NOと加工日を読む。 */
public final class InspectionSheetHeaderReader {

    private static final int MAX_ROWS = 10;
    private static final int MAX_COLS = 50;
    private static final ThreadLocal<DataFormatter> FORMATTER =
            ThreadLocal.withInitial(() -> new DataFormatter(Locale.JAPAN));

    public record Header(String iraiNo, LocalDate processingDate) {}

    private InspectionSheetHeaderReader() {}

    public static Header read(Path file) throws IOException {
        if (file == null) {
            return new Header("", null);
        }
        try {
            Header sax = readSax(file);
            if (!sax.iraiNo().isBlank() || sax.processingDate() != null) {
                return sax;
            }
        } catch (Exception ignored) {
            // ユーザモデルへフォールバック
        }
        return readUsermodel(file);
    }

    private static Header readSax(Path file) throws IOException {
        String[][] grid = new String[MAX_ROWS][MAX_COLS];
        try (OPCPackage pkg = OPCPackage.open(file.toFile(), PackageAccess.READ)) {
            XSSFReader reader = new XSSFReader(pkg);
            ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(pkg);
            StylesTable styles = reader.getStylesTable();
            XSSFReader.SheetIterator sheets = reader.getSheetIterator();
            InputStream chosen = null;
            try {
                while (sheets.hasNext()) {
                    InputStream is = sheets.next();
                    String name = sheets.getSheetName();
                    if ("検査表".equals(name)) {
                        if (chosen != null) {
                            chosen.close();
                        }
                        chosen = is;
                        while (sheets.hasNext()) {
                            sheets.next().close();
                        }
                        break;
                    }
                    if (chosen == null) {
                        chosen = is;
                    } else {
                        is.close();
                    }
                }
                if (chosen == null) {
                    return fallbackFromFileName(file);
                }
                parseSheetHeader(chosen, styles, strings, grid);
            } finally {
                if (chosen != null) {
                    chosen.close();
                }
            }
        } catch (OpenXML4JException | SAXException e) {
            throw new IOException(e.getMessage(), e);
        }
        return headerFromGrid(grid, fileName(file));
    }

    private static void parseSheetHeader(
            InputStream sheetStream, StylesTable styles, ReadOnlySharedStringsTable strings, String[][] grid)
            throws IOException {
        DataFormatter formatter = FORMATTER.get();
        XSSFSheetXMLHandler.SheetContentsHandler handler =
                new XSSFSheetXMLHandler.SheetContentsHandler() {
                    private int currentRow = -1;

                    @Override
                    public void startRow(int rowNum) {
                        if (rowNum >= MAX_ROWS) {
                            throw new HeaderRowLimit();
                        }
                        currentRow = rowNum;
                    }

                    @Override
                    public void endRow(int rowNum) {
                        currentRow = -1;
                    }

                    @Override
                    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
                        if (currentRow < 0 || currentRow >= MAX_ROWS) {
                            return;
                        }
                        int col = columnIndex(cellReference);
                        if (col < 0 || col >= MAX_COLS) {
                            return;
                        }
                        grid[currentRow][col] = formattedValue != null ? formattedValue : "";
                    }
                };
        try {
            XMLReader xmlReader = SAXHelper.newXMLReader();
            xmlReader.setContentHandler(
                    new XSSFSheetXMLHandler(styles, strings, handler, formatter, false));
            xmlReader.parse(new InputSource(sheetStream));
        } catch (HeaderRowLimit ignored) {
            // 先頭行だけ読めればよい
        } catch (ParserConfigurationException | SAXException e) {
            if (e.getCause() instanceof HeaderRowLimit) {
                return;
            }
            throw new IOException(e.getMessage(), e);
        } catch (RuntimeException e) {
            if (e instanceof HeaderRowLimit || e.getCause() instanceof HeaderRowLimit) {
                return;
            }
            throw e;
        }
    }

    private static Header readUsermodel(Path file) throws IOException {
        try (Workbook wb = PoiWorkbookOpener.open(file.toFile())) {
            Sheet sheet = selectSheet(wb);
            if (sheet == null) {
                return fallbackFromFileName(file);
            }
            String[][] grid = new String[MAX_ROWS][MAX_COLS];
            for (int r = 0; r < MAX_ROWS; r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    continue;
                }
                for (int c = 0; c < MAX_COLS; c++) {
                    Cell cell = row.getCell(c);
                    if (cell == null) {
                        continue;
                    }
                    if (cell.getCellType() == org.apache.poi.ss.usermodel.CellType.NUMERIC) {
                        Optional<LocalDate> serial =
                                InspectionSheetProcessingDate.fromExcelSerial(cell.getNumericCellValue());
                        if (serial.isPresent()) {
                            grid[r][c] = serial.get().toString();
                            continue;
                        }
                    }
                    grid[r][c] = cellText(cell);
                }
            }
            return headerFromGrid(grid, fileName(file));
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

    static Header headerFromGrid(String[][] grid, String fileName) {
        String irai = "";
        LocalDate date = null;
        if (grid != null) {
            int rows = Math.min(MAX_ROWS, grid.length);
            for (int r = 0; r < rows; r++) {
                String[] row = grid[r];
                if (row == null) {
                    continue;
                }
                int cols = Math.min(MAX_COLS, row.length);
                for (int c = 0; c < cols; c++) {
                    String text = row[c];
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
        }
        if (irai.isEmpty()) {
            irai = InspectionSheetIraiNo.extractFromFileName(fileName).orElse("");
        }
        return new Header(irai, date);
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

    private static String firstIraiToTheRight(String[] row, int labelCol) {
        if (row == null) {
            return "";
        }
        for (int c = labelCol + 1; c < Math.min(MAX_COLS, row.length); c++) {
            Optional<String> irai = InspectionSheetIraiNo.extractFromCellText(row[c]);
            if (irai.isPresent()) {
                return irai.get();
            }
        }
        return "";
    }

    private static Optional<LocalDate> dateToTheRight(String[] row, int labelCol) {
        if (row == null) {
            return Optional.empty();
        }
        for (int c = labelCol + 1; c < Math.min(MAX_COLS, row.length); c++) {
            String text = row[c];
            if (text == null || text.isBlank()) {
                continue;
            }
            Optional<LocalDate> parsed = InspectionSheetProcessingDate.parse(text);
            if (parsed.isPresent()) {
                return parsed;
            }
            Optional<LocalDate> serial = parseSerial(text);
            if (serial.isPresent()) {
                return serial;
            }
        }
        return Optional.empty();
    }

    private static Optional<LocalDate> parseSerial(String text) {
        try {
            return InspectionSheetProcessingDate.fromExcelSerial(Double.parseDouble(text.strip()));
        } catch (NumberFormatException ex) {
            return Optional.empty();
        }
    }

    private static String cellText(Cell cell) {
        if (cell == null) {
            return "";
        }
        return FORMATTER.get().formatCellValue(cell);
    }

    static int columnIndex(String cellReference) {
        if (cellReference == null || cellReference.isEmpty()) {
            return 0;
        }
        int col = 0;
        int len = cellReference.length();
        int i = cellReference.charAt(0) == '$' ? 1 : 0;
        while (i < len) {
            char ch = cellReference.charAt(i);
            if (ch >= 'A' && ch <= 'Z') {
                col = col * 26 + (ch - 'A' + 1);
            } else if (ch >= 'a' && ch <= 'z') {
                col = col * 26 + (ch - 'a' + 1);
            } else {
                break;
            }
            i++;
        }
        return Math.max(0, col - 1);
    }

    private static final class HeaderRowLimit extends RuntimeException {
        private HeaderRowLimit() {
            super("header-row-limit", null, false, false);
        }
    }
}
