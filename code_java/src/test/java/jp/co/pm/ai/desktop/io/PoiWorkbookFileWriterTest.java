package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class PoiWorkbookFileWriterTest {

    @TempDir
    Path tmp;

    @BeforeEach
    void isolateStagingRoot() {
        System.setProperty("pm.ai.test.juchuWriteStagingRoot", tmp.resolve("staging").toString());
    }

    @AfterEach
    void clearStagingRootProperty() {
        System.clearProperty("pm.ai.test.juchuWriteStagingRoot");
    }

    private static java.util.Map<String, String> ui(Path repoRoot) {
        return java.util.Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, repoRoot.toString());
    }

    @Test
    void writeReplacing_replacesTargetOnSuccess(@TempDir Path repo) throws Exception {
        Path target = repo.resolve("juchu.xlsm");
        Files.writeString(target, "old");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("受注ﾌｧｲﾙ");
            sheet.createRow(0).createCell(0).setCellValue("new");
            PoiWorkbookFileWriter.writeReplacing(target, wb, ui(repo));
        }

        assertTrue(Files.size(target) > 0L);
        try (XSSFWorkbook read = new XSSFWorkbook(Files.newInputStream(target))) {
            assertEquals("new", read.getSheet("受注ﾌｧｲﾙ").getRow(0).getCell(0).getStringCellValue());
        }
    }
}
