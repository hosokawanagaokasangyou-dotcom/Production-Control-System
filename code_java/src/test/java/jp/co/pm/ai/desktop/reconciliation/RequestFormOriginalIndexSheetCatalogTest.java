package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class RequestFormOriginalIndexSheetCatalogTest {

    @Test
    void loadAll_readsIndexRowsFromWorkbooks(@TempDir Path temp) throws Exception {
        Path originalDir = temp.resolve("original");
        Files.createDirectories(originalDir);

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet index = wb.createSheet("目次");
            index.createRow(1).createCell(0).setCellValue("加工依頼NO");
            var row = index.createRow(2);
            row.createCell(0).setCellValue("Y8-69");
            row.createCell(9).setCellValue("8/1");
            row.createCell(13).setCellValue("ABC-123");
            Path xlsm = originalDir.resolve("Y-8月（2026年）加工依頼書（国分Y）.xlsm");
            try (var out = Files.newOutputStream(xlsm)) {
                wb.write(out);
            }
        }

        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR, originalDir.toString());
        List<String> warnings = new ArrayList<>();
        List<RequestFormOriginalIndexSheetCatalog.Row> rows =
                RequestFormOriginalIndexSheetCatalog.loadAll(ui, warnings);

        assertTrue(warnings.isEmpty(), warnings.toString());
        assertEquals(1, rows.size());
        assertEquals("Y8-69", rows.getFirst().iraiNo());
        assertEquals("8/1", rows.getFirst().inputDate());
        assertEquals("ABC-123", rows.getFirst().contractNo());
        assertTrue(rows.getFirst().sourceFileName().contains("Y-8月"));
    }
}
