package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class MasterDispatchSheetsSeederTest {

    @TempDir Path tmp;

    @Test
    void importsWhenJsonMissingAndSkipsWhenJsonExists() throws Exception {
        Path json = tmp.resolve("master-dispatch-sheets.json");
        Path xlsx = tmp.resolve("master.xlsx");
        writeMinimalMaster(xlsx);

        MasterDispatchSheetsSeeder.Result first =
                MasterDispatchSheetsSeeder.loadOrImport(json, xlsx, "KONAN");
        assertEquals(MasterDispatchSheetsSeeder.Outcome.IMPORTED, first.outcome());
        assertTrue(Files.isRegularFile(json));
        assertEquals("巻返し", first.document().sheet("skills").rows().get(0).get(1));

        MasterDispatchSheetsDocument mutated =
                new MasterDispatchSheetsDocument(
                        1,
                        "KONAN",
                        json.toString(),
                        "edited",
                        java.util.Map.of(
                                "skills",
                                new MasterDispatchSheetsDocument.SheetGrid(
                                        "skills",
                                        java.util.List.of(java.util.List.of("手編集")))));
        MasterDispatchSheetsJsonStore.write(json, mutated);

        MasterDispatchSheetsSeeder.Result second =
                MasterDispatchSheetsSeeder.loadOrImport(json, xlsx, "KONAN");
        assertEquals(MasterDispatchSheetsSeeder.Outcome.LOADED_EXISTING, second.outcome());
        assertEquals("手編集", second.document().sheet("skills").rows().get(0).get(0));
    }

    @Test
    void missingSourceReturnsEmptyWithoutWritingJson() throws Exception {
        Path json = tmp.resolve("out.json");
        Path missing = tmp.resolve("no-master.xlsx");
        MasterDispatchSheetsSeeder.Result r =
                MasterDispatchSheetsSeeder.loadOrImport(json, missing, "KOKUBU");
        assertEquals(MasterDispatchSheetsSeeder.Outcome.EMPTY_MISSING_SOURCE, r.outcome());
        assertFalse(Files.exists(json));
        assertTrue(r.document().sheet("skills").rows().isEmpty());
        assertEquals("KOKUBU", r.document().factorySite());
    }

    @Test
    void reimportOverwritesExistingJson() throws Exception {
        Path json = tmp.resolve("master-dispatch-sheets.json");
        Path xlsx = tmp.resolve("master.xlsx");
        writeMinimalMaster(xlsx);
        MasterDispatchSheetsJsonStore.write(
                json, MasterDispatchSheetsDocument.empty("KONAN"));
        MasterDispatchSheetsSeeder.Result r =
                MasterDispatchSheetsSeeder.loadOrImport(json, xlsx, "KONAN", true);
        assertEquals(MasterDispatchSheetsSeeder.Outcome.IMPORTED, r.outcome());
        assertEquals("巻返し", r.document().sheet("skills").rows().get(0).get(1));
    }

    private static void writeMinimalMaster(Path master) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var skills = wb.createSheet("skills");
            skills.createRow(0).createCell(0).setCellValue("工程名");
            skills.getRow(0).createCell(1).setCellValue("巻返し");
            try (OutputStream os = Files.newOutputStream(master)) {
                wb.write(os);
            }
        }
    }
}
