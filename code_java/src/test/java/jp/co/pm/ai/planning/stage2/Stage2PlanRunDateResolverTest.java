package jp.co.pm.ai.planning.stage2;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class Stage2PlanRunDateResolverTest {

    @TempDir Path temp;

    @Test
    void resolvePlanDate_readsDataExtractionTimeFromWorkbook() throws Exception {
        Path workbook = temp.resolve("plan.xlsx");
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("加工計画DATA");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("データ抽出時間");
            sheet.createRow(1).createCell(0).setCellValue("2026/06/09 08:30:00");
            try (var out = Files.newOutputStream(workbook)) {
                wb.write(out);
            }
        }

        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK, workbook.toString());

        assertEquals(LocalDate.of(2026, 6, 9), Stage2PlanRunDateResolver.resolvePlanDate(ui));
        assertEquals("2026/06/09", Stage2PlanRunDateResolver.planDateColumnKey(ui));
    }

    @Test
    void extractDataExtractionDateTime_fallsBackToDataExtractionDateColumn() throws Exception {
        Path workbook = temp.resolve("plan2.xlsx");
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("加工計画DATA");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("データ抽出日");
            sheet.createRow(1).createCell(0).setCellValue("2026-06-11");
            try (var out = Files.newOutputStream(workbook)) {
                wb.write(out);
            }
        }

        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK, workbook.toString());

        assertTrue(
                Stage2PlanRunDateResolver.extractDataExtractionDateTime(ui)
                        .map(LocalDateTime::toLocalDate)
                        .filter(d -> d.equals(LocalDate.of(2026, 6, 11)))
                        .isPresent());
    }
}
