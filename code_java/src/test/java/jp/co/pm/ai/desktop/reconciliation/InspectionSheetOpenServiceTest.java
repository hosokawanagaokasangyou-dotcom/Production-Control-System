package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;

class InspectionSheetOpenServiceTest {

    private String priorHome;
    private String priorUserHome;

    @BeforeEach
    void setUp(@TempDir Path tmp) {
        priorHome = AppPaths.desktopAppHomeDirName();
        priorUserHome = System.getProperty("user.home");
        System.setProperty("user.home", tmp.toString());
        AppPaths.setDesktopAppHomeDirName(".pm-ai-desktop-test");
        GlobalInitSettingTarget.save(FactorySite.KONAN);
    }

    @AfterEach
    void tearDown() {
        AppPaths.setDesktopAppHomeDirName(priorHome);
        System.setProperty("user.home", priorUserHome);
    }

    @Test
    void rebuildAndFind(@TempDir Path root) throws Exception {
        Path month = root.resolve("2026年").resolve("9月");
        Files.createDirectories(month);
        Path xlsx = month.resolve("2026_C8-9(SEC済)完了.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sh = wb.createSheet("検査表");
            var r2 = sh.createRow(1);
            r2.createCell(26).setCellValue("加工日");
            r2.createCell(28).setCellValue(46261);
            var r3 = sh.createRow(2);
            r3.createCell(0).setCellValue("加工依頼№");
            r3.createCell(3).setCellValue("C8-9");
            try (var out = Files.newOutputStream(xlsx)) {
                wb.write(out);
            }
        }
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_INSPECTION_SHEET_DIR, root.toString());
        InspectionSheetOpenService.RebuildResult rebuilt = InspectionSheetOpenService.rebuild(ui, null);
        assertEquals(1, rebuilt.rows().size());
        List<InspectionSheetIndexStore.Row> hits = InspectionSheetOpenService.find(ui, "c8-9");
        assertEquals(1, hits.size());
        assertEquals("C8-9", hits.get(0).iraiNo());
        assertTrue(Files.isRegularFile(InspectionSheetIndexStore.indexFile(FactorySite.KONAN)));
    }
}
