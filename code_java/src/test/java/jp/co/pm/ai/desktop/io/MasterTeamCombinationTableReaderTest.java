package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class MasterTeamCombinationTableReaderTest {

    @TempDir Path tmp;

    @Test
    void readsComboKeysFromTeamCombinationSheet() throws Exception {
        Path master = tmp.resolve("master.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sh = wb.createSheet(MasterTeamCombinationTableReader.SHEET_NAME);
            var h = sh.createRow(0);
            h.createCell(0).setCellValue("組み合わせ行ID");
            h.createCell(1).setCellValue("工程名");
            h.createCell(2).setCellValue("機械名");
            var r1 = sh.createRow(1);
            r1.createCell(1).setCellValue("巻返し");
            r1.createCell(2).setCellValue("フィルム挿入機(間紙)");
            try (OutputStream os = Files.newOutputStream(master)) {
                wb.write(os);
            }
        }
        Set<String> keys = MasterTeamCombinationTableReader.readNormalizedComboKeys(master);
        assertEquals(1, keys.size());
        assertTrue(keys.contains("巻返し+フィルム挿入機(間紙)"));
    }

    @Test
    void emptyWhenSheetMissing() throws Exception {
        Path master = tmp.resolve("master.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("skills");
            try (OutputStream os = Files.newOutputStream(master)) {
                wb.write(os);
            }
        }
        assertTrue(MasterTeamCombinationTableReader.readNormalizedComboKeys(master).isEmpty());
    }
}
