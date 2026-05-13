package jp.co.pm.ai.planning.stage2.parity;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class Stage2WorkbookSemanticParityTest {

    @Test
    void identicalBooks(@TempDir Path root) throws Exception {
        Path a = root.resolve("a.xlsx");
        Path b = root.resolve("b.xlsx");
        writeOneSheet(a, "T", new String[] {"c1", "c2"}, new String[][] {{"a", "b"}});
        writeOneSheet(b, "T", new String[] {"c1", "c2"}, new String[][] {{"a", "b"}});
        Stage2ParityCheckResult r = Stage2WorkbookSemanticParity.compareXlsx(a, b);
        assertTrue(r.identical());
    }

    @Test
    void differentCell(@TempDir Path root) throws Exception {
        Path a = root.resolve("a.xlsx");
        Path b = root.resolve("b.xlsx");
        writeOneSheet(a, "T", new String[] {"c1"}, new String[][] {{"x"}});
        writeOneSheet(b, "T", new String[] {"c1"}, new String[][] {{"y"}});
        Stage2ParityCheckResult r = Stage2WorkbookSemanticParity.compareXlsx(a, b);
        assertFalse(r.identical());
    }

    private static void writeOneSheet(Path path, String sheetName, String[] headers, String[][] rows)
            throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sh = wb.createSheet(sheetName);
            Row hr = sh.createRow(0);
            for (int c = 0; c < headers.length; c++) {
                hr.createCell(c).setCellValue(headers[c]);
            }
            for (int r = 0; r < rows.length; r++) {
                Row rr = sh.createRow(r + 1);
                for (int c = 0; c < rows[r].length; c++) {
                    rr.createCell(c).setCellValue(rows[r][c]);
                }
            }
            Files.createDirectories(path.getParent());
            try (var os = Files.newOutputStream(path)) {
                wb.write(os);
            }
        }
    }
}
