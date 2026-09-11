package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class InspectionSheetIndexScannerTest {

    @Test
    void scansRecursivelyAndSkipsTempAndCsv(@TempDir Path tmp) throws Exception {
        Path month = tmp.resolve("2026年").resolve("9月");
        Files.createDirectories(month);
        Path xlsx = month.resolve("2026_C8-9(SEC済)完了.xlsx");
        writeKonan(xlsx, "C8-9", 46261);
        Files.writeString(month.resolve("ignore.csv"), "a,b");
        Files.writeString(month.resolve("~$lock.xlsx"), "x");

        InspectionSheetIndexScanner.Result first =
                InspectionSheetIndexScanner.scan(tmp, List.of(), null);
        assertEquals(1, first.rows().size());
        assertEquals("C8-9", first.rows().get(0).iraiNo());
        assertEquals(1, first.readExcelCount());

        InspectionSheetIndexScanner.Result second =
                InspectionSheetIndexScanner.scan(tmp, first.rows(), null);
        assertEquals(1, second.rows().size());
        assertEquals(0, second.readExcelCount());
    }

    private static void writeKonan(Path file, String irai, double serial) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sh = wb.createSheet("検査表");
            var r2 = sh.createRow(1);
            r2.createCell(26).setCellValue("加工日");
            r2.createCell(28).setCellValue(serial);
            var r3 = sh.createRow(2);
            r3.createCell(0).setCellValue("加工依頼№");
            r3.createCell(3).setCellValue(irai);
            try (var out = Files.newOutputStream(file)) {
                wb.write(out);
            }
        }
    }
}
