package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Reads Excel sheet {@code \u8a2d\u5b9a_\u74b0\u5883\u5909\u6570} (same row rules as {@code workbook_env_bootstrap.py}).
 */
public final class WorkbookEnvSheetReader {

    public static final String SHEET_NAME = "\u8a2d\u5b9a_\u74b0\u5883\u5909\u6570";

    public record RowEntry(String key, String value, String description) {}

    private WorkbookEnvSheetReader() {}

    public static List<RowEntry> read(Path workbookPath) throws IOException {
        Objects.requireNonNull(workbookPath, "workbookPath");
        if (!Files.isRegularFile(workbookPath)) {
            throw new IOException("not a file: " + workbookPath);
        }
        DataFormatter fmt = new DataFormatter(Locale.ROOT);
        try (InputStream in = Files.newInputStream(workbookPath);
                Workbook wb = WorkbookFactory.create(in)) {
            Sheet sh = wb.getSheet(SHEET_NAME);
            if (sh == null) {
                throw new IOException("sheet not found: " + SHEET_NAME + " in " + workbookPath);
            }
            List<RowEntry> out = new ArrayList<>();
            int last = sh.getLastRowNum();
            List<Row> rows = new ArrayList<>();
            for (int i = 0; i <= last; i++) {
                Row row = sh.getRow(i);
                if (row != null) {
                    rows.add(row);
                }
            }
            if (rows.isEmpty()) {
                return out;
            }
            int start = 0;
            Row head = rows.get(0);
            String hk = cellStr(fmt, head.getCell(0)).toLowerCase(Locale.ROOT);
            if (hk.equals("\u5909\u6570\u540d")
                    || hk.equals("name")
                    || hk.equals("key")
                    || hk.equals("\u74b0\u5883\u5909\u6570")
                    || hk.equals("env")) {
                start = 1;
            }
            for (int ri = start; ri < rows.size(); ri++) {
                Row row = rows.get(ri);
                String k = cellStr(fmt, row.getCell(0));
                if (k.isEmpty()) {
                    continue;
                }
                if (k.startsWith("#")) {
                    continue;
                }
                String v = cellStr(fmt, row.getCell(1));
                String desc = cellStr(fmt, row.getCell(2));
                out.add(new RowEntry(k, v, desc.isEmpty() ? null : desc));
            }
            return out;
        }
    }

    private static String cellStr(DataFormatter fmt, Cell cell) {
        if (cell == null) {
            return "";
        }
        return fmt.formatCellValue(cell).trim();
    }
}
