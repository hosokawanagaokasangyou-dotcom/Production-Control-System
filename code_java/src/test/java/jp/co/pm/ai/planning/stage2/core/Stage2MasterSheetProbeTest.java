package jp.co.pm.ai.planning.stage2.core;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class Stage2MasterSheetProbeTest {

    @Test
    void scan_detectsNeedAndCalendarSheets(@TempDir Path tmp) throws Exception {
        Path master = tmp.resolve("m.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("skills");
            wb.createSheet("need");
            wb.createSheet("機械カレンダー");
            try (OutputStream os = Files.newOutputStream(master)) {
                wb.write(os);
            }
        }
        Stage2MasterSheetProbe p = Stage2MasterSheetProbe.scan(master);
        assertEquals(3, p.sheetCount());
        assertTrue(p.hasNeedSheet());
        assertTrue(p.hasMachineCalendarSheet());
    }

    @Test
    void scan_minimalMasterFromSmoke(@TempDir Path tmp) throws Exception {
        Path master = tmp.resolve("master.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("skills");
            wb.createSheet("メイン");
            try (OutputStream os = Files.newOutputStream(master)) {
                wb.write(os);
            }
        }
        Stage2MasterSheetProbe p = Stage2MasterSheetProbe.scan(master);
        assertEquals(2, p.sheetCount());
        assertFalse(p.hasNeedSheet());
        assertFalse(p.hasMachineCalendarSheet());
    }
}
