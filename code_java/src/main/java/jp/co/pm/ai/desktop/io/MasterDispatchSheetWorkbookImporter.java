package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.OffsetDateTime;
import java.time.ZoneId;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * 現在工場の master ブックから 4 シートを格子として読む。他工場のファイルは開かない。
 */
public final class MasterDispatchSheetWorkbookImporter {

    private static final ZoneId TOKYO = ZoneId.of("Asia/Tokyo");

    private MasterDispatchSheetWorkbookImporter() {}

    public static MasterDispatchSheetsDocument importWorkbook(Path workbookPath, String factorySite)
            throws IOException {
        Objects.requireNonNull(workbookPath, "workbookPath");
        if (!Files.isRegularFile(workbookPath)) {
            throw new IOException("not a file: " + workbookPath);
        }
        DataFormatter fmt = new DataFormatter(Locale.ROOT);
        Map<String, MasterDispatchSheetsDocument.SheetGrid> sheets = new LinkedHashMap<>();
        try (InputStream in = Files.newInputStream(workbookPath);
                Workbook wb = WorkbookFactory.create(in)) {
            for (String key : MasterDispatchSheetsDocument.SHEET_KEYS) {
                String excelName = MasterDispatchSheetsDocument.defaultSheetName(key);
                Sheet sh = wb.getSheet(excelName);
                sheets.put(key, new MasterDispatchSheetsDocument.SheetGrid(excelName, readTrimmed(sh, fmt)));
            }
        }
        String importedAt = OffsetDateTime.now(TOKYO).truncatedTo(ChronoUnit.SECONDS).toString();
        return new MasterDispatchSheetsDocument(
                MasterDispatchSheetsDocument.SCHEMA_VERSION,
                factorySite != null ? factorySite : "",
                workbookPath.toAbsolutePath().normalize().toString(),
                importedAt,
                sheets);
    }

    static List<List<String>> readTrimmed(Sheet sh, DataFormatter fmt) {
        if (sh == null) {
            return List.of();
        }
        int lastRow = sh.getLastRowNum();
        if (lastRow < 0) {
            return List.of();
        }
        int maxCol = 0;
        List<List<String>> raw = new ArrayList<>();
        for (int r = 0; r <= lastRow; r++) {
            Row row = sh.getRow(r);
            int lastCell = row != null ? row.getLastCellNum() : -1;
            if (lastCell > maxCol) {
                maxCol = lastCell;
            }
            List<String> cells = new ArrayList<>();
            if (row != null && lastCell > 0) {
                for (int c = 0; c < lastCell; c++) {
                    cells.add(cellStr(fmt, row.getCell(c)));
                }
            }
            raw.add(cells);
        }
        int lastUsedRow = -1;
        int lastUsedCol = -1;
        for (int r = 0; r < raw.size(); r++) {
            List<String> row = raw.get(r);
            for (int c = 0; c < row.size(); c++) {
                if (!row.get(c).isEmpty()) {
                    lastUsedRow = r;
                    if (c > lastUsedCol) {
                        lastUsedCol = c;
                    }
                }
            }
        }
        if (lastUsedRow < 0) {
            return List.of();
        }
        int width = lastUsedCol + 1;
        List<List<String>> out = new ArrayList<>(lastUsedRow + 1);
        for (int r = 0; r <= lastUsedRow; r++) {
            List<String> row = raw.get(r);
            List<String> padded = new ArrayList<>(width);
            for (int c = 0; c < width; c++) {
                padded.add(c < row.size() ? row.get(c) : "");
            }
            out.add(List.copyOf(padded));
        }
        return List.copyOf(out);
    }

    private static String cellStr(DataFormatter fmt, Cell cell) {
        if (cell == null) {
            return "";
        }
        String s = fmt.formatCellValue(cell);
        return s != null ? s.strip() : "";
    }
}
